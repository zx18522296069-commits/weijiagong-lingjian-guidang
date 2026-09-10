"""未加工零件云端自动处理主入口。

正式流程：
1. 只扫描“正在加工”订单源（支持文件夹快捷方式）
2. 读取永久累计台账
3. 只扫描“拆图结果”根目录直接的 *_完成.xlsx
4. 按完整板材号去重并逐板核验
5. 更新累计台账与当前待加工零件
6. 回读验证成功后，才把已成功入账的拆图结果移入“已录入数量”

注意：整单完成只控制“当前待加工零件”动态视图是否移除订单，
不控制拆图结果文件归档；每张板成功入账后即可归档。
"""

from __future__ import annotations

import os
import tempfile
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

from modules.drive_manager import DriveManager
from modules.excel_generator import generate_cumulative_report, generate_pending_report
from modules.excel_reader import read_existing_ledger, read_source_summary, read_split_result
from modules.idempotency import reconcile_posted_board
from modules.logger import get_logger
from modules.process_parts import (
    build_current_state,
    build_source_index,
    validate_new_board,
)

logger = get_logger()
BEIJING_TZ = timezone(timedelta(hours=8))


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"缺少环境变量 {name}")
    return value


def _row_key(row: dict) -> tuple:
    return (
        str(row.get("order", "")).strip(),
        str(row.get("drawing", "")).strip(),
        float(row.get("thickness", 0)),
        str(row.get("bevel", "")).strip(),
    )


def _anomaly_signature(row: dict) -> tuple:
    return (
        str(row.get("板材号", "")),
        str(row.get("异常类型", "")),
        str(row.get("说明", "")),
    )


def _append_unique_anomaly(target: list[dict], seen: set[tuple], row: dict) -> bool:
    sig = _anomaly_signature(row)
    if sig in seen:
        return False
    target.append(row)
    seen.add(sig)
    return True


def _blocked_anomaly(board_id: str, reason: str, quantity: int = 0) -> dict:
    return {
        "板材号": board_id,
        "业务日期": datetime.now(BEIJING_TZ).date().isoformat(),
        "订单号": "",
        "图号": "",
        "厚度(mm)": "",
        "数量": quantity,
        "异常类型": "阻断入账",
        "说明": f"{reason}；未重复累计、未归档，文件保留在拆图结果根目录。",
    }


def _verify_output_workbooks(cumulative_path: Path, pending_path: Path):
    cumulative = load_workbook(cumulative_path, read_only=True, data_only=True)
    expected = {"累计加工台账", "加工流水", "板材入账记录", "异常记录"}
    if not expected.issubset(set(cumulative.sheetnames)):
        raise RuntimeError(f"累计台账生成后缺少工作表: {expected - set(cumulative.sheetnames)}")
    pending = load_workbook(pending_path, read_only=True, data_only=True)
    if "当前待加工零件" not in pending.sheetnames:
        raise RuntimeError("当前待加工零件文件缺少正式工作表")


def _split_historical_rows(existing: dict, source_index: dict) -> tuple[list[dict], list[dict]]:
    """
    累计台账永久保留历史：
    - 已从“正在加工”退出且历史剩余=0的零件，继续保留在累计台账。
    - 若源文件消失时仍存在非0剩余（含负数异常），阻断整次运行，不能把历史当成已结束。
    """
    completed_history: list[dict] = []
    missing_incomplete: list[dict] = []
    for row in existing.get("historical_rows", []):
        if _row_key(row) in source_index:
            continue
        if int(row.get("remaining", 0)) == 0:
            completed_history.append(row)
        else:
            missing_incomplete.append(row)
    return completed_history, missing_incomplete


def run():
    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    mode_name = "只读测试" if test_mode else "正式生产"
    logger.info(f"开始执行未加工零件自动更新任务｜模式={mode_name}")

    cumulative_file_id = _required_env("CUMULATIVE_FILE_ID")
    pending_file_id = _required_env("PENDING_FILE_ID")

    drive = DriveManager()
    if not drive.check_config():
        raise RuntimeError("Google Drive 配置不完整")

    source_files = drive.list_order_source_files()
    pending_files = drive.list_pending_split_files()
    logger.info(f"发现订单原始汇总表 {len(source_files)} 个")
    logger.info(f"拆图结果根目录待处理完成文件 {len(pending_files)} 个")

    with tempfile.TemporaryDirectory() as temp_dir_str:
        temp_dir = Path(temp_dir_str)

        # 1) 下载并读取当前永久累计台账。
        current_ledger_path = temp_dir / "累计加工台账_当前.xlsx"
        drive.download_file(cumulative_file_id, str(current_ledger_path))
        existing = read_existing_ledger(current_ledger_path)
        logger.info(f"永久入账板材数 {len(existing['posted_boards'])}")

        # 2) 读取所有当前订单原始汇总表。任何源文件结构异常都停止正式输出。
        source_records: list[dict] = []
        for idx, item in enumerate(source_files, start=1):
            local = temp_dir / f"source_{idx}_{Path(item['name']).name}"
            drive.download_file(item["id"], str(local))
            records = read_source_summary(
                local,
                source_name=item["name"],
                order_container_name=item.get("order_container_name", ""),
            )
            source_records.extend(records)
            logger.info(f"读取订单源 {item.get('order_container_name') or item['name']}：{len(records)} 条零件")

        source_index = build_source_index(source_records)
        logger.info(f"当前订单原始零件总条数 {len(source_records)}")

        # 累计台账永久性保护：已完成退出订单保留；未完成却丢失源文件则阻断。
        completed_history, missing_incomplete = _split_historical_rows(existing, source_index)
        if missing_incomplete:
            affected_orders = sorted({row["order"] for row in missing_incomplete})
            examples = ", ".join(
                f"{row['order']}/{row['drawing']}/剩余{row['remaining']}"
                for row in missing_incomplete[:5]
            )
            raise RuntimeError(
                "发现累计台账中仍未完成/异常的零件已不在“正在加工”原始订单源中，"
                f"禁止静默删除。受影响订单={affected_orders}；示例={examples}"
            )
        if completed_history:
            logger.info(f"永久累计台账保留已退出的完成历史零件 {len(completed_history)} 条")

        # 3) 逐板预校验；整板任一行失败则整板阻断。
        delta_by_key: dict[tuple, int] = defaultdict(int)
        source_additions: dict[tuple, list[tuple[str, int]]] = defaultdict(list)
        new_flows: list[dict] = []
        new_board_records: list[dict] = []
        new_anomalies: list[dict] = []
        accepted_files: list[dict] = []
        reconciled_files: list[dict] = []
        blocked: list[tuple[str, str]] = []

        anomaly_seen = {_anomaly_signature(r) for r in existing["anomalies"]}
        posted_boards = set(existing["posted_boards"])

        for idx, item in enumerate(sorted(pending_files, key=lambda x: x.get("name", "")), start=1):
            filename = item["name"]
            board_id = drive.board_id_from_filename(filename)
            local = temp_dir / f"pending_{idx}_{Path(filename).name}"
            drive.download_file(item["id"], str(local))
            payload = read_split_result(local, board_id=board_id, source_name=filename)

            # 上次可能已经成功写入台账，但在“移动到已录入数量”之前中断。
            # 这种情况不能再次累计；只有根目录文件与永久台账完全一致时，才允许补归档。
            if board_id in posted_boards:
                recovery = reconcile_posted_board(
                    board_id=board_id,
                    filename=filename,
                    split_payload=payload,
                    source_records=source_records,
                    existing_flows=existing["flows"],
                    existing_board_records=existing["board_records"],
                )
                if recovery["ok"]:
                    reconciled_files.append({**item, "board_id": board_id})
                    logger.info(f"板材 {board_id} 已入账且内容一致：本轮只补归档，不重复累计")
                else:
                    reason = recovery["reason"]
                    blocked.append((board_id, reason))
                    qty = sum(int(r.get("split_quantity", 0)) for r in payload.get("rows", []))
                    _append_unique_anomaly(new_anomalies, anomaly_seen, _blocked_anomaly(board_id, reason, qty))
                    logger.warning(f"板材 {board_id} 已入账但根目录文件无法安全补归档：{reason}")
                continue

            result = validate_new_board(
                board_id=board_id,
                filename=filename,
                split_payload=payload,
                source_records=source_records,
                posted_boards=posted_boards,
            )

            if not result["ok"]:
                blocked.append((board_id, result["error"]))
                _append_unique_anomaly(new_anomalies, anomaly_seen, result["anomaly"])
                logger.warning(f"板材 {board_id} 阻断：{result['error']}")
                continue

            for key, qty in result["deltas"].items():
                delta_by_key[key] += qty
            for key, qty in result["board_sources"].items():
                source_additions[key].append((board_id, qty))
            new_flows.extend(result["flows"])
            new_board_records.append(result["board_record"])
            _append_unique_anomaly(new_anomalies, anomaly_seen, result["anomaly"])
            accepted_files.append({**item, "board_id": board_id})
            posted_boards.add(board_id)
            logger.info(f"板材 {board_id} 校验通过：计入 {result['board_record']['计入件数']} 件")

        # 4) 以当前原始资料 + 既有累计 + 本次增量生成活动订单状态。
        active_state_rows = build_current_state(
            source_records,
            existing_state=existing["state"],
            delta_by_key=dict(delta_by_key),
            source_additions=dict(source_additions),
        )
        cumulative_state_rows = active_state_rows + completed_history

        total_remaining = sum(int(r["remaining"]) for r in active_state_rows)
        total_pending_weight = sum(float(r["pending_weight_t"]) for r in active_state_rows)
        logger.info(f"生成后活动订单当前剩余件数 {total_remaining}")
        logger.info(f"生成后活动订单当前未出重量 {total_pending_weight:.3f} t")

        run_note = (
            f"自动执行：扫描订单源{len(source_files)}个、待处理板材{len(pending_files)}张；"
            f"本次新入账{len(accepted_files)}张、补归档候选{len(reconciled_files)}张、阻断{len(blocked)}张；"
            f"永久保留已退出完成历史{len(completed_history)}条。"
            "所有尺寸、订单数量、基础总重量均以正在加工目录原始汇总表为事实源；"
            "拆图结果成功写入并回读验证后才归档。"
        )

        all_flows = existing["flows"] + new_flows
        all_board_records = existing["board_records"] + new_board_records
        all_anomalies = existing["anomalies"] + new_anomalies

        cumulative_out = temp_dir / "累计加工台账.xlsx"
        pending_out = temp_dir / "当前待加工零件.xlsx"
        generate_cumulative_report(
            cumulative_state_rows,
            all_flows,
            all_board_records,
            all_anomalies,
            cumulative_out,
            note=run_note,
        )
        generate_pending_report(active_state_rows, pending_out, note=run_note)
        _verify_output_workbooks(cumulative_out, pending_out)

        if test_mode:
            logger.info("只读测试完成：未写回 Drive、未移动任何拆图结果文件")
            if reconciled_files:
                logger.info(f"测试发现 {len(reconciled_files)} 张已入账但尚未归档的文件，可在正式模式安全补归档")
            if blocked:
                for board_id, reason in blocked:
                    logger.warning(f"测试发现阻断板材 {board_id}: {reason}")
            return

        # 5) 正式模式：覆盖两个固定正式文件 ID，避免产生同名副本。
        drive.update_file_content(cumulative_file_id, str(cumulative_out))
        drive.update_file_content(pending_file_id, str(pending_out))
        logger.info("两个正式 Excel 已写回 Drive，开始回读验证")

        verify_cumulative = temp_dir / "verify_累计加工台账.xlsx"
        verify_pending = temp_dir / "verify_当前待加工零件.xlsx"
        drive.download_file(cumulative_file_id, str(verify_cumulative))
        drive.download_file(pending_file_id, str(verify_pending))
        _verify_output_workbooks(verify_cumulative, verify_pending)
        verified_ledger = read_existing_ledger(verify_cumulative)
        verify_board_ids = [item["board_id"] for item in accepted_files + reconciled_files]
        missing_boards = [board_id for board_id in verify_board_ids if board_id not in verified_ledger["posted_boards"]]
        if missing_boards:
            raise RuntimeError(f"写回后板材入账记录验证失败: {missing_boards}")
        logger.info("正式文件回读验证通过")

        # 6) 新入账板材 + 已入账但上次归档中断的板材，在回读成功后统一移动。
        for item in accepted_files + reconciled_files:
            drive.move_file(item["id"], drive.archive_folder_id)
            logger.info(f"已归档拆图结果: {item['name']} -> 拆图结果/已录入数量")

        if blocked:
            for board_id, reason in blocked:
                logger.warning(f"阻断板材仍保留根目录 {board_id}: {reason}")

        logger.info("正式生产任务执行结束")


if __name__ == "__main__":
    run()
