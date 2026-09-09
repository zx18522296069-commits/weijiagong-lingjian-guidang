"""生产重试/中断恢复：识别“已入账但文件尚未归档”的同一张板。"""

from __future__ import annotations

from collections import defaultdict

from modules.process_parts import validate_new_board


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _flow_key(row: dict) -> tuple:
    return (
        _text(row.get("图号")),
        float(row.get("厚度(mm)") or 0),
        _text(row.get("坡口")),
    )


def _aggregate_flows(rows: list[dict], board_id: str) -> dict[tuple, int]:
    result: dict[tuple, int] = defaultdict(int)
    for row in rows:
        if _text(row.get("板材号")) != board_id:
            continue
        try:
            qty = int(round(float(row.get("本张板加工数量") or 0)))
        except (TypeError, ValueError):
            continue
        result[_flow_key(row)] += qty
    return dict(result)


def reconcile_posted_board(
    *,
    board_id: str,
    filename: str,
    split_payload: dict,
    source_records: list[dict],
    existing_flows: list[dict],
    existing_board_records: list[dict],
) -> dict:
    """
    如果板材号已在永久台账中，但原完成文件仍留在拆图结果根目录，
    仅在“当前文件内容按原始订单重新校验通过”且“永久台账净流水/入账件数与该文件一致”时，
    才判定为上次写入后归档步骤中断，可在本轮回读验证后补归档。
    否则阻断，绝不再次累计。
    """
    expected = validate_new_board(
        board_id=board_id,
        filename=filename,
        split_payload=split_payload,
        source_records=source_records,
        posted_boards=set(),
    )
    if not expected["ok"]:
        return {"ok": False, "reason": f"已入账板材的根目录文件重新校验失败：{expected['error']}"}

    expected_flows: dict[tuple, int] = defaultdict(int)
    for row in expected["flows"]:
        expected_flows[_flow_key(row)] += int(row["本张板加工数量"])
    expected_flows = dict(expected_flows)
    actual_flows = _aggregate_flows(existing_flows, board_id)
    if actual_flows != expected_flows:
        return {
            "ok": False,
            "reason": f"已入账板材与永久加工流水不一致：文件={expected_flows}，台账净流水={actual_flows}",
        }

    board_records = [
        row for row in existing_board_records
        if _text(row.get("板材号")) == board_id
        and _text(row.get("状态")) not in {"作废", "未入账", "阻断"}
    ]
    if not board_records:
        return {"ok": False, "reason": "板材号在去重集合中但没有可验证的正式入账记录"}

    latest = board_records[-1]
    try:
        ledger_qty = int(round(float(latest.get("计入件数") or 0)))
    except (TypeError, ValueError):
        return {"ok": False, "reason": "永久板材入账记录的计入件数无效"}
    expected_qty = int(expected["board_record"]["计入件数"])
    if ledger_qty != expected_qty:
        return {
            "ok": False,
            "reason": f"已入账板材计入件数与根目录文件不一致：台账={ledger_qty}，文件={expected_qty}",
        }

    return {
        "ok": True,
        "reason": "永久台账与根目录完成文件一致，判定为已入账但未完成归档的重试场景",
        "expected_qty": expected_qty,
    }
