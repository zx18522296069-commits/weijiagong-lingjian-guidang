"""未加工零件云端自动处理主入口。

正式流程：
1. 扫描“正在加工”订单源，按 metadata 签名增量读取
2. 在“正在加工”根目录按文件名发现两份正式表；缺失时安全自愈
3. 读取永久累计台账（若缺失则以当前全部订单、累计已加工=0作为新基线）
4. 只扫描“拆图结果”根目录直接的 *_完成.xlsx/xlsm/xls
5. 按“完整板材号 + 文件业务内容”去重并逐板整板核验
6. 批量更新累计事实并完整重建两份正式 Excel
7. 本地校验后按顺序上传；每份上传只用 Drive metadata/MD5/size 确认，不重新下载整个 Excel
8. 两份正式文件都确认成功后，才把新入账和补归档文件统一移入“已录入数量”

注意：整单完成同时控制《当前待加工零件》的退出，以及《累计加工台账.xlsx》中
“累计加工台账”主页面向“已完成订单”页的迁移；加工流水、板材入账记录、异常记录永久保留。
整单完成不控制拆图结果文件归档；每张板成功入账并完成事务确认后即可归档。
"""

from __future__ import annotations

import hashlib
import os
import tempfile
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

from modules.drive_manager import DriveManager
from modules.excel_generator import generate_cumulative_report, generate_pending_report
from modules.excel_reader import (
    normalize_order_key,
    read_existing_ledger,
    read_source_summary,
    read_split_result,
)
from modules.formal_files import (
    CUMULATIVE_NAME,
    PENDING_NAME,
    discover_formal_files,
    empty_ledger,
    write_formal_file,
)
from modules.idempotency import reconcile_posted_board
from modules.logger import get_logger
from modules.process_parts import (
    board_content_fingerprint,
    build_current_state,
    build_source_index,
    validate_new_board,
)
from modules.source_cache import SourceRecordCache

logger = get_logger()
BEIJING_TZ = timezone(timedelta(hours=8))


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
        "处理结论": "未移动到已录入数量",
        "未移动原因": reason,
        "处理建议": "补齐或修正“正在加工”中的对应订单原始汇总表；确认图号、厚度、基础件数和基础总重量后，下一次会自动重新核验。",
        "说明": f"{reason}；未重复累计、未归档，文件保留在拆图结果根目录。",
    }


def _verify_output_workbooks(cumulative_path: Path, pending_path: Path):
    cumulative = load_workbook(cumulative_path, read_only=True, data_only=True)
    expected = {"累计加工台账", "加工流水", "板材入账记录", "异常记录", "已完成订单"}
    if not expected.issubset(set(cumulative.sheetnames)):
        raise RuntimeError(f"累计台账生成后缺少工作表: {expected - set(cumulative.sheetnames)}")
    pending = load_workbook(pending_path, read_only=True, data_only=True)
    if "当前待加工零件" not in pending.sheetnames:
        raise RuntimeError("当前待加工零件文件缺少正式工作表")


def _file_md5(path: Path) -> str:
    digest = hashlib.md5()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _verify_drive_upload(drive: DriveManager, file_id: str, local_path: Path) -> dict:
    metadata = drive.get_file(file_id, fields="id,name,modifiedTime,md5Checksum,size")
    local_md5 = _file_md5(local_path)
    remote_md5 = str(metadata.get("md5Checksum") or "").lower()
    if not remote_md5:
        raise RuntimeError(f"Drive 上传验证失败：{local_path.name} 未返回 md5Checksum")
    if remote_md5 != local_md5:
        raise RuntimeError(
            f"Drive 上传验证失败：{local_path.name} MD5 不一致，本地={local_md5}，远端={remote_md5}"
        )
    remote_size = metadata.get("size")
    if remote_size not in (None, "") and int(remote_size) != local_path.stat().st_size:
        raise RuntimeError(
            f"Drive 上传验证失败：{local_path.name} size 不一致，本地={local_path.stat().st_size}，远端={remote_size}"
        )
    return metadata


def _split_historical_rows(existing: dict, source_index: dict) -> tuple[list[dict], list[dict]]:
    """保留已完成历史；源文件消失但仍未完成的历史行冻结保留。"""
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


def _read_current_sources(drive: DriveManager, source_files: list[dict], temp_dir: Path):
    """增量读取当前全部订单源；缓存只影响速度，不改变事实源。"""
    source_cache_file = Path(os.getenv("SOURCE_RECORD_CACHE_FILE", ".cache/source_records.json"))
    source_cache = SourceRecordCache(source_cache_file)
    source_records: list[dict] = []
    source_errors: list[dict] = []
    cached_sources = 0
    refreshed_sources = 0

    for idx, item in enumerate(source_files, start=1):
        container_name = item.get("order_container_name", "")
        cached = source_cache.get(item)
        if cached:
            cached_sources += 1
            if cached["status"] == "error":
                error_item = dict(cached["error"])
                source_errors.append(error_item)
                logger.warning(
                    f"复用订单源异常缓存 {error_item.get('order') or error_item.get('source')}: "
                    f"{error_item.get('reason')}"
                )
                continue
            records = list(cached["records"])
            source_records.extend(records)
            logger.info(f"复用订单源缓存 {container_name or item['name']}：{len(records)} 条零件")
            continue

        refreshed_sources += 1
        local = temp_dir / f"source_{idx}_{Path(item['name']).name}"
        drive.download_file(item["id"], str(local))
        try:
            records = read_source_summary(
                local,
                source_name=item["name"],
                order_container_name=container_name,
            )
        except ValueError as error:
            order = normalize_order_key(container_name)
            error_item = {
                "order": order,
                "source": container_name or item["name"],
                "reason": str(error),
            }
            source_errors.append(error_item)
            source_cache.store_error(item, error_item)
            logger.warning(f"订单源未完成 {order or container_name or item['name']}: {error}")
            continue
        source_records.extend(records)
        source_cache.store_records(item, records)
        logger.info(f"重新读取订单源 {container_name or item['name']}：{len(records)} 条零件")

    source_cache.retain(source_files)
    try:
        source_cache.save()
    except OSError as error:
        logger.warning(f"订单源增量缓存保存失败，本次结果仍按真实源文件执行：{error}")

    logger.info(f"订单源增量处理：复用缓存 {cached_sources} 个，重新下载解析 {refreshed_sources} 个")
    return source_records, source_errors, cached_sources, refreshed_sources


def run():
    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    mode_name = "只读测试" if test_mode else "正式生产"
    logger.info(f"开始执行未加工零件自动更新任务｜模式={mode_name}")

    drive = DriveManager()
    if not drive.check_config():
        raise RuntimeError("Google Drive 配置不完整")

    source_files = drive.list_order_source_files()
    pending_files = drive.list_pending_split_files()
    formal_files = discover_formal_files(drive)
    cumulative_meta = formal_files["cumulative"]
    pending_meta = formal_files["pending"]
    cumulative_missing = cumulative_meta is None
    pending_missing = pending_meta is None

    logger.info(f"发现订单原始汇总表 {len(source_files)} 个")
    logger.info(f"拆图结果根目录待处理完成文件 {len(pending_files)} 个")
    logger.info(
        f"正式文件检查：{CUMULATIVE_NAME}={'缺失' if cumulative_missing else '存在'}；"
        f"{PENDING_NAME}={'缺失' if pending_missing else '存在'}"
    )

    if not source_files and (cumulative_missing or pending_missing):
        raise RuntimeError("正式台账缺失，但“正在加工”中没有可用订单原始汇总表，无法安全重建")

    with tempfile.TemporaryDirectory() as temp_dir_str:
        temp_dir = Path(temp_dir_str)

        # 1) 先读取“正在加工”全部当前订单。重建正式表时必须以完整订单集为事实源。
        source_records, source_errors, cached_sources, refreshed_sources = _read_current_sources(
            drive, source_files, temp_dir
        )
        source_index = build_source_index(source_records)
        logger.info(f"当前订单原始零件总条数 {len(source_records)}")

        if (cumulative_missing or pending_missing) and source_errors:
            details = "；".join(
                f"{item.get('order') or item.get('source')}：{item.get('reason')}"
                for item in source_errors[:5]
            )
            raise RuntimeError(
                "正式台账缺失且存在不可读取订单源，禁止用不完整订单数据重建。"
                f"请先修正订单源：{details}"
            )

        # 2) 累计台账存在则读取永久历史；缺失则建立空历史基线。
        if cumulative_meta:
            current_ledger_path = temp_dir / "累计加工台账_当前.xlsx"
            drive.download_file(cumulative_meta["id"], str(current_ledger_path))
            existing = read_existing_ledger(current_ledger_path)
            logger.info(f"永久入账板材数 {len(existing['posted_boards'])}")
        else:
            existing = empty_ledger()
            logger.warning(
                f"{CUMULATIVE_NAME} 缺失：将以当前全部订单重新建立，历史累计从 0 开始；"
                "本次拆图结果根目录中的完成文件仍会按正常规则核验并计入。"
            )

        # 3) 累计台账永久性保护：已完成退出订单保留；异常/缺源订单冻结。
        unavailable_orders = {item["order"] for item in source_errors if item["order"]}
        frozen_source_rows = [
            row for row in existing.get("historical_rows", [])
            if row.get("order") in unavailable_orders
        ]
        history_for_check = {
            **existing,
            "historical_rows": [
                row for row in existing.get("historical_rows", [])
                if row.get("order") not in unavailable_orders
            ],
        }
        completed_history, missing_incomplete = _split_historical_rows(history_for_check, source_index)
        if missing_incomplete:
            affected_orders = sorted({row["order"] for row in missing_incomplete})
            examples = ", ".join(
                f"{row['order']}/{row['drawing']}/剩余{row['remaining']}"
                for row in missing_incomplete[:5]
            )
            frozen_source_rows.extend(missing_incomplete)
            logger.warning(
                "历史订单源缺失，相关零件已冻结并保留上次结果；继续处理其余板材。"
                f"受影响订单={affected_orders}；示例={examples}"
            )
        if completed_history:
            logger.info(f"永久累计台账保留已退出的完成历史零件 {len(completed_history)} 条")
        if frozen_source_rows:
            logger.warning(f"因订单源异常冻结历史数据 {len(frozen_source_rows)} 条，保持上次结果不变")

        # 4) 逐板预校验；整板任一行失败则整板阻断。
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
        posted_board_keys = set(existing.get("posted_board_keys", set()))
        legacy_posted_boards = set(existing.get("legacy_posted_boards", set()))

        logger.info("已启用补归档规则：已入账且内容核验一致的根目录板材只归档，不重复累计或扣减")
        for idx, item in enumerate(sorted(pending_files, key=lambda x: x.get("name", "")), start=1):
            filename = item["name"]
            board_id = drive.board_id_from_filename(filename)
            local = temp_dir / f"pending_{idx}_{Path(filename).name}"
            drive.download_file(item["id"], str(local))
            payload = read_split_result(local, board_id=board_id, source_name=filename)
            content_fingerprint = board_content_fingerprint(payload)

            exact_duplicate = (board_id, content_fingerprint) in posted_board_keys
            legacy_candidate = board_id in legacy_posted_boards
            if exact_duplicate:
                reconciled_files.append({**item, "board_id": board_id, "content_fingerprint": content_fingerprint})
                logger.info(f"板材 {board_id} 重复内容已核验，将在两份正式表确认成功后补归档")
                continue

            if legacy_candidate:
                recovery = reconcile_posted_board(
                    board_id=board_id,
                    filename=filename,
                    split_payload=payload,
                    source_records=source_records,
                    existing_flows=existing["flows"] + new_flows,
                    existing_board_records=existing["board_records"] + new_board_records,
                    content_fingerprint="",
                )
                if recovery["ok"]:
                    reconciled_files.append({**item, "board_id": board_id, "content_fingerprint": content_fingerprint})
                    logger.info(f"板材 {board_id} 与历史台账内容一致，将在两份正式表确认成功后补归档")
                    continue
                if recovery.get("comparable") is False:
                    proof = recovery.get("historical_proof", {})
                    reason = (
                        f"检测到历史入账记录：完整板材号={board_id}，"
                        f"原文件={proof.get('filename') or '未知'}，已入账={proof.get('quantity') or '未知'}件，"
                        f"状态={proof.get('status') or '未知'}；"
                        f"本次无法复核内容，不重复扣减、不归档。原因：{recovery['reason']}"
                    )
                    blocked.append((board_id, reason))
                    qty = sum(int(r.get("split_quantity", 0)) for r in payload.get("rows", []))
                    _append_unique_anomaly(new_anomalies, anomaly_seen, _blocked_anomaly(board_id, reason, qty))
                    logger.warning(f"板材 {board_id} 已有历史入账记录但本次无法复核，已保留根目录")
                    continue
                logger.info(f"板材 {board_id} 编号重复且内容不同：进入同板材号内容冲突校验，禁止自动再次入账")

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
            accepted_files.append({**item, "board_id": board_id, "content_fingerprint": content_fingerprint})
            posted_boards.add(board_id)
            posted_board_keys.add((board_id, content_fingerprint))
            logger.info(f"板材 {board_id} 校验通过：计入 {result['board_record']['计入件数']} 件")

        # 5) 当前原始订单 + 既有累计 + 本次增量，重建两份正式表的内容。
        active_state_rows = build_current_state(
            source_records,
            existing_state=existing["state"],
            delta_by_key=dict(delta_by_key),
            source_additions=dict(source_additions),
        )
        active_state_rows.extend(frozen_source_rows)
        cumulative_state_rows = active_state_rows + completed_history

        total_remaining = sum(int(r["remaining"]) for r in active_state_rows)
        total_pending_weight = sum(float(r["pending_weight_t"]) for r in active_state_rows)
        logger.info(f"生成后活动订单当前剩余件数 {total_remaining}")
        logger.info(f"生成后活动订单当前未出重量 {total_pending_weight:.3f} t")

        rebuild_note = []
        if cumulative_missing:
            rebuild_note.append(f"重建{CUMULATIVE_NAME}")
        if pending_missing:
            rebuild_note.append(f"重建{PENDING_NAME}")
        run_note = (
            f"自动执行：扫描订单源{len(source_files)}个、缓存复用{cached_sources}个、重新读取{refreshed_sources}个、"
            f"成功可用{len(source_files) - len(source_errors)}个、未完成{len(source_errors)}个、待处理板材{len(pending_files)}张；"
            f"本次新入账{len(accepted_files)}张、补归档候选{len(reconciled_files)}张、阻断{len(blocked)}张；"
            f"永久保留已退出完成历史{len(completed_history)}条。"
            + (f"自愈：{'、'.join(rebuild_note)}。" if rebuild_note else "")
            + "所有尺寸、订单数量、基础总重量均以正在加工目录原始汇总表为事实源；"
            "拆图结果只有在两份正式表本地校验且上传文件通过Drive metadata/MD5/size确认后才归档。"
        )

        all_flows = existing["flows"] + new_flows
        all_board_records = existing["board_records"] + new_board_records
        all_anomalies = existing["anomalies"] + new_anomalies

        cumulative_out = temp_dir / CUMULATIVE_NAME
        pending_out = temp_dir / PENDING_NAME
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
            if rebuild_note:
                logger.warning(f"只读测试发现正式文件缺失，将在正式运行时自动{'、'.join(rebuild_note)}")
            logger.info("只读测试完成：未写回 Drive、未移动任何拆图结果文件")
            if source_errors:
                for item in source_errors:
                    logger.warning(f"未完成订单 {item['order'] or item['source']}: {item['reason']}")
            if reconciled_files:
                logger.info(f"测试发现 {len(reconciled_files)} 张已入账但尚未归档的文件，可在正式模式安全补归档")
            if blocked:
                for board_id, reason in blocked:
                    logger.warning(f"测试发现阻断板材 {board_id}: {reason}")
            return

        # 6) 可恢复事务：累计表先写并校验成功，再写当前待加工表。
        cumulative_written = write_formal_file(drive, cumulative_meta, cumulative_out)
        cumulative_file_id = cumulative_written["id"]
        if not cumulative_written.get("_skipped"):
            _verify_drive_upload(drive, cumulative_file_id, cumulative_out)
            logger.info(f"上传校验通过｜{CUMULATIVE_NAME}｜Drive metadata/MD5/size 一致")

        pending_written = write_formal_file(drive, pending_meta, pending_out)
        pending_file_id = pending_written["id"]
        if not pending_written.get("_skipped"):
            _verify_drive_upload(drive, pending_file_id, pending_out)
            logger.info(f"上传校验通过｜{PENDING_NAME}｜Drive metadata/MD5/size 一致")

        if cumulative_written.get("_skipped") and pending_written.get("_skipped"):
            logger.info("正式文件写回：业务事实无变化，两份正式 Excel 均未上传")
        else:
            logger.info(
                f"正式文件写回：{CUMULATIVE_NAME}={'跳过' if cumulative_written.get('_skipped') else ('新建' if cumulative_missing else '更新')}；"
                f"{PENDING_NAME}={'跳过' if pending_written.get('_skipped') else ('新建' if pending_missing else '更新')}"
            )

        # 7) 远端字节已由 MD5/size 确认与本地一致；语义校验直接使用本地已确认文件，禁止再次下载整个 Excel。
        verified_ledger = read_existing_ledger(cumulative_out)
        missing_boards = [
            item["board_id"]
            for item in accepted_files
            if (item["board_id"], item["content_fingerprint"]) not in verified_ledger.get("posted_board_keys", set())
        ]
        if missing_boards:
            raise RuntimeError(f"写回后板材入账记录验证失败: {missing_boards}")
        logger.info("正式文件验证通过：本地语义校验完成；远端上传使用 Drive metadata/MD5/size 确认，未重新下载整个 Excel")

        # 8) 两份正式文件都成功后，统一归档新入账和已确认重复文件。
        for item in accepted_files:
            drive.move_file(item["id"], drive.archive_folder_id)
            logger.info(f"已归档拆图结果: {item['name']} -> 拆图结果/已录入数量")
            logger.info(
                f"板材处理结果｜文件={item['name']}｜板材={item['board_id']}｜"
                "状态=已累计、已录入｜原因=首次校验通过并已写回累计台账｜处理建议=无"
            )
        for item in reconciled_files:
            drive.move_file(item["id"], drive.archive_folder_id)
            logger.info(f"已归档拆图结果: {item['name']} -> 拆图结果/已录入数量")
            logger.info(
                f"板材处理结果｜文件={item['name']}｜板材={item['board_id']}｜"
                "状态=已累计、仅补归档｜原因=累计台账已有相同内容记录，本次未重复累计｜处理建议=无"
            )

        if blocked:
            for board_id, reason in blocked:
                logger.warning(f"阻断板材仍保留根目录 {board_id}: {reason}")
                logger.warning(
                    f"板材处理结果｜文件={board_id}_完成｜板材={board_id}｜"
                    f"状态=未累计、未记录｜原因={reason}｜"
                    "处理建议=核对该板拆图结果和对应订单原始汇总表后重新执行。"
                )

        logger.info("正式生产任务执行结束")


if __name__ == "__main__":
    run()
