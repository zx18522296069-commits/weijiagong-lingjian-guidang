"""未加工零件正式业务逻辑。只做确定性匹配，不做模糊猜测。"""

from __future__ import annotations

import math
import hashlib
import json
from collections import defaultdict
from datetime import datetime, timezone, timedelta

BEIJING_TZ = timezone(timedelta(hours=8))


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _thickness(value) -> float:
    return float(value)


def part_key(order: str, drawing: str, thickness, bevel: str) -> tuple:
    return (_text(order), _text(drawing), _thickness(thickness), _text(bevel))


def board_content_fingerprint(split_payload: dict) -> str:
    """以规范化业务内容识别同一张板；与文件名、行顺序无关。"""
    rows = []
    for row in split_payload.get("rows", []):
        rows.append({
            "order": _text(row.get("order")),
            "drawing": _text(row.get("drawing")),
            "thickness": float(row.get("thickness") or 0),
            "bevel": _text(row.get("bevel")),
            "base_quantity": int(row.get("base_quantity") or 0),
            "split_quantity": int(row.get("split_quantity") or 0),
            "base_total_weight_t": round(float(row.get("base_total_weight_t") or 0), 9),
        })
    canonical = json.dumps(
        sorted(rows, key=lambda item: json.dumps(item, ensure_ascii=False, sort_keys=True)),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:16]


def build_source_index(source_records: list[dict]) -> dict[tuple, dict]:
    index: dict[tuple, dict] = {}
    for row in source_records:
        key = part_key(row["order"], row["drawing"], row["thickness"], row["bevel"])
        if key in index:
            raise ValueError(f"原始汇总表存在重复业务键，不能唯一处理: {key}")
        index[key] = row
    return index


def _same_thickness(a, b) -> bool:
    return math.isclose(float(a), float(b), rel_tol=0.0, abs_tol=1e-9)


def _drawing_candidates(source_records: list[dict], split: dict) -> tuple[list[dict], str]:
    """
    正式匹配：订单+图号+厚度(+坡口)。
    唯一允许的图号规范化：拆图结果省略 D53K 型号前缀时，以 `.后缀` 唯一匹配。
    """
    base = [
        row for row in source_records
        if row["order"] == split["order"]
        and _same_thickness(row["thickness"], split["thickness"])
    ]

    exact = [row for row in base if row["drawing"] == split["drawing"]]
    mode = "exact"
    candidates = exact

    if not candidates and split["order"].upper().startswith("D53K-"):
        suffix = split["drawing"]
        candidates = [
            row for row in base
            if row["drawing"].endswith("." + suffix)
        ]
        mode = "d53k_suffix"

    if split.get("bevel"):
        candidates = [row for row in candidates if row["bevel"] == split["bevel"]]
        mode += "+bevel"

    return candidates, mode


def match_split_row(source_records: list[dict], split: dict) -> tuple[dict, str]:
    candidates, mode = _drawing_candidates(source_records, split)
    if not candidates:
        raise ValueError(
            f"无匹配：订单={split['order']} 图号={split['drawing']} 厚度={split['thickness']}"
            + (f" 坡口={split['bevel']}" if split.get("bevel") else "")
        )
    if len(candidates) != 1:
        raise ValueError(
            f"匹配不唯一：订单={split['order']} 图号={split['drawing']} 厚度={split['thickness']}，候选={len(candidates)}"
        )

    source = candidates[0]
    if split["base_quantity"] != source["quantity"]:
        raise ValueError(
            f"基础数量不一致：{split['order']} {split['drawing']}，拆图={split['base_quantity']}，原始={source['quantity']}"
        )
    if not math.isclose(
        float(split["base_total_weight_t"]),
        float(source["total_weight_t"]),
        rel_tol=1e-9,
        abs_tol=1e-6,
    ):
        raise ValueError(
            f"基础总重量不一致：{split['order']} {split['drawing']}，拆图={split['base_total_weight_t']}t，原始={source['total_weight_t']}t"
        )
    return source, mode


def _append_board_source(existing: str, board_id: str, quantity: int) -> str:
    token = f"{board_id}×{quantity}"
    existing = _text(existing)
    return token if not existing else f"{existing}；{token}"


def validate_new_board(
    *,
    board_id: str,
    filename: str,
    split_payload: dict,
    source_records: list[dict],
    posted_boards: set[str],
) -> dict:
    """
    对一张新板做全量预校验。任何一行失败，则整张板阻断，不产生正式入账流水。
    """
    today = datetime.now(BEIJING_TZ).date().isoformat()
    # 板材号允许重复；真正的去重键是“板材号 + 内容指纹”，由主流程在
    # 全量校验前判断。保留 posted_boards 参数仅兼容既有调用。
    content_fingerprint = board_content_fingerprint(split_payload)

    matched = []
    try:
        for split in split_payload["rows"]:
            source, mode = match_split_row(source_records, split)
            matched.append((split, source, mode))
    except Exception as exc:
        return {
            "ok": False,
            "error": str(exc),
            "anomaly": {
                "板材号": board_id,
                "业务日期": today,
                "订单号": split_payload.get("rows", [{}])[0].get("order", ""),
                "图号": "",
                "厚度(mm)": "",
                "数量": sum(int(r.get("split_quantity", 0)) for r in split_payload.get("rows", [])),
                "异常类型": "阻断入账",
                "说明": f"{exc}；未入账、未归档，文件保留在拆图结果根目录。",
            },
        }

    deltas: dict[tuple, int] = defaultdict(int)
    board_sources: dict[tuple, int] = defaultdict(int)
    flows: list[dict] = []
    for split, source, mode in matched:
        key = part_key(source["order"], source["drawing"], source["thickness"], source["bevel"])
        qty = int(split["split_quantity"])
        deltas[key] += qty
        board_sources[key] += qty
        note_parts = [
            f"业务日期{today}",
            f"订单{source['order']}",
            "基础数量与基础总重量与当前订单原始汇总表核对通过",
        ]
        if not split.get("bevel"):
            note_parts.append("拆图结果未提供坡口列，按订单号+图号+厚度唯一匹配，坡口取原始汇总表")
        if mode.startswith("d53k_suffix"):
            note_parts.append(f"图号按唯一规范化由{split['drawing']}匹配{source['drawing']}")
        flows.append({
            "板材号": board_id,
            "内容指纹": content_fingerprint,
            "拆图结果文件": filename,
            "图号": source["drawing"],
            "厚度(mm)": source["thickness"],
            "坡口": source["bevel"],
            "本张板加工数量": qty,
            "数据性质": "正式入账",
            "说明": "；".join(note_parts) + "。",
        })

    total_qty = sum(deltas.values())
    review = split_payload.get("review") or {}
    abs_diff = review.get("abs_diff_kg")
    if abs_diff not in (None, ""):
        abs_diff = float(abs_diff)
        anomaly_type = "无异常" if abs_diff <= 1.0 else "复核差异（不阻断）"
        review_note = f"图片重量复核差额绝对值{abs_diff:.2f}kg；图片重量仅用于复核，不用于反推或分摊。"
    else:
        anomaly_type = "复核信息缺失（不阻断）"
        review_note = "拆图结果未读取到图片重量差额；不影响基础数量/基础总重量正式校验。"

    orders = sorted({source["order"] for _, source, _ in matched})
    anomaly = {
        "板材号": board_id,
        "业务日期": today,
        "订单号": "；".join(orders),
        "图号": "",
        "厚度(mm)": "",
        "数量": total_qty,
        "异常类型": anomaly_type,
        "说明": f"{len(matched)}条记录、{total_qty}件；基础数量与基础总重量均核对通过。{review_note}",
    }

    board_record = {
        "数据性质": "正式入账",
        "板材号": board_id,
        "内容指纹": content_fingerprint,
        "拆图结果文件": filename,
        "计入件数": total_qty,
        "状态": "已核验并入账",
        "归档位置": "拆图结果/已录入数量",
    }

    return {
        "ok": True,
        "deltas": dict(deltas),
        "board_sources": dict(board_sources),
        "flows": flows,
        "board_record": board_record,
        "content_fingerprint": content_fingerprint,
        "anomaly": anomaly,
    }


def build_current_state(
    source_records: list[dict],
    existing_state: dict[tuple, dict],
    delta_by_key: dict[tuple, int] | None = None,
    source_additions: dict[tuple, list[tuple[str, int]]] | None = None,
) -> list[dict]:
    delta_by_key = delta_by_key or {}
    source_additions = source_additions or {}
    rows = []

    for source in source_records:
        key = part_key(source["order"], source["drawing"], source["thickness"], source["bevel"])
        old = existing_state.get(key, {})
        processed = int(old.get("processed", 0)) + int(delta_by_key.get(key, 0))
        remaining = int(source["quantity"]) - processed
        board_sources = _text(old.get("board_sources", ""))
        for board_id, qty in source_additions.get(key, []):
            board_sources = _append_board_source(board_sources, board_id, qty)

        unit_weight = float(source["total_weight_t"]) / int(source["quantity"])
        pending_weight = unit_weight * max(remaining, 0)
        if remaining < 0:
            status = "超加工/待核查"
        elif remaining == 0:
            status = "已完成"
        elif processed > 0:
            status = "部分完成"
        else:
            status = "未开始"

        rows.append({
            **source,
            "board_sources": board_sources,
            "processed": processed,
            "remaining": remaining,
            "pending_weight_t": pending_weight,
            "status": status,
        })
    return rows


def group_by_order(rows: list[dict]) -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        grouped[row["order"]].append(row)
    return dict(grouped)


def order_is_complete(rows: list[dict]) -> bool:
    """仅用于“当前待加工零件”动态视图移除；不控制拆图结果文件归档。"""
    return bool(rows) and all(int(row["remaining"]) == 0 for row in rows)


def order_has_anomaly(rows: list[dict]) -> bool:
    return any(int(row["remaining"]) < 0 for row in rows)
