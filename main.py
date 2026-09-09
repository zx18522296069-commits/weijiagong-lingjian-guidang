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
import shutil
import tempfile
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from modules.drive_manager import DriveManager
from modules.excel_generator import generate_cumulative_report, generate_pending_report
from modules.excel_reader import read_existing_ledger, read_source_summary, read_split_result
from modules.logger import get_logger
from modules.process_parts import (
    build_current_state,
    build_source_index,
    validate_new_board,
)

logger = get_logger()


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"缺少环境变量 {name}")
    return value


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


def _verify_output_workbooks(cumulative_path: Path, pending_path: Path):
    cumulative = load_workbook(cumulative_path, read_only=True, data_only=True)
    expected = {"累计加工台账", "加工流水", "板材入账记录", "异常记录"}
    if not expected.issubset(set(cumulative.sheetnames)):
        raise RuntimeError(f"累计台账生成后缺少工作表: {expected - set(cumulative.sheetnames)}")
    pending = load_workbook(pending_path, read_only=True, data_only=True)
    if "当前待加工零件" not in pending.sheetnames:
        raise RuntimeError("当前待加工零件文件缺少正式工作表")


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

    artifact_dir = Path("artifacts")
    artifact_dir.mkdir(parents=True, exist_ok=True)

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

        build_source_index(source_records)  # 强制检查正式业务键唯一。
        logger.info(f"当前订单原始零件总条数 {len(source_records)}")

        # 3) 逐板预校验；整板任一行失败则整板阻断。
        delta_by_key: dict[tuple, int] = defaultdict(int)
        source_additions: dict[tuple, list[tuple[str, int]]] = defaultdict(list)
        new_flows: list[dict] = []
        new_board_records: list[dict] = []
        new_anomalies: list[dict] = []
        accepted_files: list[dict] = []
        blocked: list[tuple[str, str]] = []

        anomaly_seen = {_anomaly_signature(r) for r in existing["anomalies"]}
        posted_boards = set(existing["posted_boards"])

        for idx, item in enumerate(sorted(pending_files, key=lambda x: x.get("name", "")), start=1):
            filename = item["name"]
            board_id = drive.board_id_from_filename(filename)
            local = temp_dir / f"pending_{idx}_{Path(filename).name}"
            drive.download_file(item["id"], str(local))
            payload = read_split_result(local, board_id=board_id, source_name=filename)
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

        # 4) 以订单原始资料 + 既有累计 + 本次增量生成当前状态。
        state_rows = build_current_state(
            source_records,
            existing_state=existing["state"],
            delta_by_key=dict(delta_by_key),
            source_additions=dict(source_additions),
        )

        total_remaining = sum(int(r["remaining"]) for r in state_rows)
        total_pending_weight = sum(float(r["pending_weight_t"]) for r in state_rows)
        logger.info(f"生成后当前剩余件数 {total_remaining}")
        logger.info(f"生成后当前未出重量 {total_pending_weight:.3f} t")

        run_note = (
            f"自动执行：扫描订单源{len(source_files)}个、待处理板材{len(pending_files)}张；"
            f"本次校验通过{len(accepted_files)}张、阻断{len(blocked)}张。"
            "所有尺寸、订单数量、基础总重量均以正在加工目录原始汇总表为事实源；"
            "拆图结果成功写入并回读验证后才归档。"
        )

        all_flows = existing["flows"] + new_flows
        all_board_records = existing["board_records"] + new_board_records
        all_anomalies = existing["anomalies"] + new_anomalies

        cumulative_out = temp_dir / "累计加工台账.xlsx"
        pending_out = temp_dir / "当前待加工零件.xlsx"
        generate_cumulative_report(
            state_rows,
            all_flows,
            all_board_records,
            all_anomalies,
            cumulative_out,
            note=run_note,
        )
        generate_pending_report(state_rows, pending_out, note=run_note)
        _verify_output_workbooks(cumulative_out, pending_out)

        # 测试模式也保留本次生成物为 GitHub Actions artifact，便于核对，但绝不写 Drive。
        shutil.copy2(cumulative_out, artifact_dir / "累计加工台账_测试生成.xlsx")
        shutil.copy2(pending_out, artifact_dir / "当前待加工零件_测试生成.xlsx")

        if test_mode:
            logger.info("只读测试完成：未写回 Drive、未移动任何拆图结果文件")
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
        missing_boards = [
            item["board_id"] for item in accepted_files
            if item["board_id"] not in verified_ledger["posted_boards"]
        ]
        if missing_boards:
            raise RuntimeError(f"写回后板材入账记录验证失败: {missing_boards}")
        logger.info("正式文件回读验证通过")

        # 6) 只有经过本次全量核验、成功写入、回读确认的板材文件才移动到已录入数量。
        for item in accepted_files:
            drive.move_file(item["id"], drive.archive_folder_id)
            logger.info(f"已归档拆图结果: {item['name']} -> 拆图结果/已录入数量")

        if blocked:
            for board_id, reason in blocked:
                logger.warning(f"阻断板材仍保留根目录 {board_id}: {reason}")

        logger.info("正式生产任务执行结束")


if __name__ == "__main__":
    run()
