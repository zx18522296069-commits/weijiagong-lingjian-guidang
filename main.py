"""未加工零件：增量扫描、整板核验、批量入账、全量重建视图、hash确认后归档。"""
from __future__ import annotations

import os
import tempfile
import time
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

from modules.drive_manager import DriveManager
from modules.excel_generator import GREEN, RED, YELLOW, generate_cumulative_report, generate_pending_report
from modules.excel_reader import normalize_order_key, read_existing_ledger, read_source_summary, read_split_result
from modules.formal_files import CUMULATIVE_NAME, PENDING_NAME, discover_formal_files, empty_ledger, write_formal_file
from modules.idempotency import reconcile_posted_board
from modules.logger import get_logger
from modules.process_parts import board_content_fingerprint, build_current_state, build_source_index, validate_new_board
from modules.runtime_state import RuntimeState

logger = get_logger(); BEIJING_TZ = timezone(timedelta(hours=8))


def _row_key(row: dict) -> tuple:
    return (str(row.get("order", "")).strip(), str(row.get("drawing", "")).strip(), float(row.get("thickness", 0)), str(row.get("bevel", "")).strip())


def _anomaly_signature(row: dict) -> tuple:
    return (str(row.get("板材号", "")), str(row.get("异常类型", "")), str(row.get("说明", "")))


def _append_unique_anomaly(target, seen, row):
    sig = _anomaly_signature(row)
    if sig in seen: return False
    target.append(row); seen.add(sig); return True


def _blocked_anomaly(board_id: str, reason: str, quantity: int = 0) -> dict:
    return {"板材号": board_id, "业务日期": datetime.now(BEIJING_TZ).date().isoformat(), "订单号": "", "图号": "", "厚度(mm)": "",
            "数量": quantity, "异常类型": "阻断入账", "处理结论": "未移动到已录入数量", "未移动原因": reason,
            "处理建议": "补齐或修正对应订单原始汇总表后重新执行。", "说明": f"{reason}；未重复累计、未归档。"}


def _business_key_change_anomaly(row: dict) -> dict:
    return {"板材号": "", "业务日期": datetime.now(BEIJING_TZ).date().isoformat(), "订单号": row.get("order", ""),
            "图号": row.get("drawing", ""), "厚度(mm)": row.get("thickness", ""), "数量": row.get("processed", 0),
            "异常类型": "订单源业务键变化", "处理结论": "冻结历史", "未移动原因": "已存在历史加工量的业务键在当前订单源中消失",
            "处理建议": "人工确认图号/厚度/坡口变更后再决定历史归属。",
            "说明": "旧业务键已有累计加工记录，禁止自动迁移到新零件；历史状态冻结保留。"}


def _overprocess_anomaly(row: dict) -> dict:
    return {"板材号": "", "业务日期": datetime.now(BEIJING_TZ).date().isoformat(), "订单号": row["order"], "图号": row["drawing"],
            "厚度(mm)": row["thickness"], "数量": row["processed"], "异常类型": "超加工/待核查", "处理结论": "照实入账",
            "未移动原因": "", "处理建议": "核对订单数量和历史加工流水。",
            "说明": f"订单总数量={row['quantity']}，累计已加工={row['processed']}，当前剩余={row['remaining']}；负数未被修正。"}


def _split_historical_rows(existing: dict, source_index: dict, current_orders: set[str]):
    completed, missing, key_changed = [], [], []
    for row in existing.get("historical_rows", []):
        if _row_key(row) in source_index: continue
        if row.get("order") in current_orders and int(row.get("processed", 0)) > 0:
            key_changed.append(row)
        elif int(row.get("remaining", 0)) == 0:
            completed.append(row)
        else:
            missing.append(row)
    return completed, missing, key_changed


def _read_sources(drive: DriveManager, items: list[dict], state: RuntimeState, temp_dir: Path):
    changes = state.classify_sources(items)
    records, errors = [], []
    source_downloads = 0; parse_seconds = 0.0
    for idx, item in enumerate(items, 1):
        cached = state.get_source(item)
        if cached:
            if cached.get("status") == "error": errors.append(dict(cached["error"]))
            else: records.extend(cached.get("records", []))
            continue
        local = temp_dir / f"source_{idx}_{Path(item['name']).name}"
        t = time.perf_counter(); drive.download_file(item["id"], str(local)); source_downloads += 1
        try:
            p = time.perf_counter(); parsed = read_source_summary(local, source_name=item["name"], order_container_name=item.get("order_container_name", "")); parse_seconds += time.perf_counter() - p
            records.extend(parsed); state.set_source_records(item, parsed)
        except ValueError as exc:
            err = {"order": normalize_order_key(item.get("order_container_name", "")), "source": item.get("order_container_name") or item["name"], "reason": str(exc)}
            errors.append(err); state.set_source_error(item, err)
    state.retain_sources(items)
    logger.info("Drive扫描｜当前订单源=%d｜unchanged=%d｜added=%d｜modified=%d｜removed=%d｜实际下载订单Excel=%d",
                len(items), len(changes['unchanged']), len(changes['added']), len(changes['modified']), len(changes['removed']), source_downloads)
    return records, errors, changes, source_downloads, parse_seconds


def _validate_output(cumulative_path: Path, pending_path: Path, expected_rows: list[dict], accepted: list[dict]) -> dict:
    cumulative = load_workbook(cumulative_path, data_only=False)
    required = {"累计加工台账", "加工流水", "板材入账记录", "异常记录"}
    if not required.issubset(cumulative.sheetnames): raise RuntimeError(f"累计台账缺少工作表: {required - set(cumulative.sheetnames)}")
    pending = load_workbook(pending_path, data_only=False)
    if "当前待加工零件" not in pending.sheetnames: raise RuntimeError("当前待加工零件缺少正式工作表")
    parsed = read_existing_ledger(cumulative_path)
    parsed_rows = {_row_key(r): r for r in parsed["historical_rows"]}
    for row in expected_rows:
        actual = parsed_rows.get(_row_key(row))
        if actual is None: raise RuntimeError(f"本地校验缺少零件: {_row_key(row)}")
        if int(actual["processed"]) != int(row["processed"]) or int(actual["remaining"]) != int(row["remaining"]):
            raise RuntimeError(f"本地校验数量不一致: {_row_key(row)}")
    for item in accepted:
        if (item["board_id"], item["content_fingerprint"]) not in parsed.get("posted_board_keys", set()):
            raise RuntimeError(f"本地校验缺少板材入账记录: {item['board_id']}")
    # 校验主表状态颜色；黄色=未加工/部分完成，绿色=已完成，红色=超加工。
    ws = cumulative["累计加工台账"]
    expected_fill = {"未加工": YELLOW, "部分完成": YELLOW, "已完成": GREEN, "超加工/待核查": RED}
    for row in ws.iter_rows():
        values = [cell.value for cell in row]
        if values and values[-1] in expected_fill:
            wanted = expected_fill[values[-1]]
            if any((cell.fill.fgColor.rgb or "").upper().replace("00", "", 1)[-6:] != wanted for cell in row[:12]):
                raise RuntimeError(f"本地校验状态颜色不一致: {values[-1]}")
    return parsed


def run():
    started = time.perf_counter(); timings = defaultdict(float)
    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    logger.info("开始执行未加工零件自动更新任务｜模式=%s", "dry-run/只读测试" if test_mode else "正式生产")
    drive = DriveManager()
    if not drive.check_config(): raise RuntimeError("Google Drive 配置不完整")
    state_path = Path(os.getenv("RUNTIME_STATE_FILE", "runtime_state/state.json")); t = time.perf_counter(); state = RuntimeState(state_path); timings["cache恢复"] = time.perf_counter()-t

    t = time.perf_counter(); source_files = drive.list_order_source_files(); pending_files = drive.list_pending_split_files(); formal = discover_formal_files(drive); timings["metadata扫描"] = time.perf_counter()-t
    cumulative_meta, pending_meta = formal["cumulative"], formal["pending"]
    if not source_files and (not cumulative_meta or not pending_meta): raise RuntimeError("正式台账缺失且没有可用订单源")

    with tempfile.TemporaryDirectory() as td:
        temp_dir = Path(td)
        source_records, source_errors, changes, source_downloads, parse_seconds = _read_sources(drive, source_files, state, temp_dir); timings["Excel解析"] += parse_seconds
        source_index = build_source_index(source_records); current_orders = {r["order"] for r in source_records}

        ledger_downloads = 0; ledger_cache_hit = bool(cumulative_meta and state.ledger_matches(cumulative_meta))
        if ledger_cache_hit:
            existing = state.get_ledger()
            logger.info("累计台账｜cache命中｜正式累计Excel下载=0｜Drive MD5变化=否")
        elif cumulative_meta:
            p = temp_dir / "累计加工台账_当前.xlsx"; t = time.perf_counter(); drive.download_file(cumulative_meta["id"], str(p)); timings["Drive下载"] += time.perf_counter()-t; ledger_downloads = 1
            t = time.perf_counter(); existing = read_existing_ledger(p); timings["Excel解析"] += time.perf_counter()-t
            logger.info("累计台账｜cache未命中｜正式累计Excel下载=1｜Drive MD5变化=%s", "是" if state.ledger else "首次/缓存无效")
        else:
            existing = empty_ledger(); logger.warning("累计加工台账缺失：按当前订单建立空历史基线")

        if existing is None: raise RuntimeError("runtime state中的累计台账缓存损坏")
        if (not cumulative_meta or not pending_meta) and source_errors:
            raise RuntimeError("正式台账缺失且存在不可读取订单源，禁止用不完整订单数据重建")

        # 真正无变化快速退出：只做metadata扫描，不下载正文、不上传。
        if cumulative_meta and pending_meta and ledger_cache_hit and not pending_files and all(not changes[k] for k in ("added","modified","removed")):
            logger.info("拆图结果｜根目录待处理=0｜新入账=0｜重复补归档=0｜阻断=0")
            logger.info("本轮Drive正文下载文件总数=0｜订单=0｜累计台账=0｜拆图结果=0")
            logger.info("无业务变化，快速退出；正式Excel上传=0")
            logger.info("时间统计｜metadata扫描=%.3fs｜cache恢复=%.3fs｜总耗时=%.3fs", timings['metadata扫描'], timings['cache恢复'], time.perf_counter()-started)
            return

        unavailable = {x["order"] for x in source_errors if x.get("order")}
        pre_frozen = [r for r in existing.get("historical_rows", []) if r.get("order") in unavailable]
        check_existing = {**existing, "historical_rows": [r for r in existing.get("historical_rows", []) if r.get("order") not in unavailable]}
        completed_history, missing_incomplete, key_changed = _split_historical_rows(check_existing, source_index, current_orders)
        frozen_rows = pre_frozen + missing_incomplete + key_changed

        delta_by_key, source_additions = defaultdict(int), defaultdict(list)
        new_flows, new_board_records, new_anomalies = [], [], []
        accepted, reconciled, blocked = [], [], []
        anomaly_seen = {_anomaly_signature(r) for r in existing.get("anomalies", [])}
        for row in key_changed: _append_unique_anomaly(new_anomalies, anomaly_seen, _business_key_change_anomaly(row))
        posted_boards = set(existing.get("posted_boards", set())); posted_keys = set(existing.get("posted_board_keys", set())); legacy = set(existing.get("legacy_posted_boards", set()))

        pending_downloads = 0; t_verify = time.perf_counter()
        for idx, item in enumerate(sorted(pending_files, key=lambda x: x.get("name", "")), 1):
            board_id = drive.board_id_from_filename(item["name"]); local = temp_dir / f"pending_{idx}_{Path(item['name']).name}"
            t = time.perf_counter(); drive.download_file(item["id"], str(local)); timings["Drive下载"] += time.perf_counter()-t; pending_downloads += 1
            payload = read_split_result(local, board_id=board_id, source_name=item["name"]); fp = board_content_fingerprint(payload)
            if (board_id, fp) in posted_keys:
                reconciled.append({**item, "board_id": board_id, "content_fingerprint": fp}); continue
            if board_id in posted_boards and board_id not in legacy:
                reason = "同板材号内容冲突：该完整板材号已入账，但当前文件内容指纹不同"
                blocked.append((board_id, reason)); _append_unique_anomaly(new_anomalies, anomaly_seen, _blocked_anomaly(board_id, reason)); continue
            if board_id in legacy:
                recovery = reconcile_posted_board(board_id=board_id, filename=item["name"], split_payload=payload, source_records=source_records,
                    existing_flows=existing["flows"]+new_flows, existing_board_records=existing["board_records"]+new_board_records, content_fingerprint="")
                if recovery.get("ok"):
                    reconciled.append({**item, "board_id": board_id, "content_fingerprint": fp}); continue
                reason = f"同板材号内容冲突：历史记录无法与当前文件唯一一致复核；{recovery.get('reason','')}"
                blocked.append((board_id, reason)); _append_unique_anomaly(new_anomalies, anomaly_seen, _blocked_anomaly(board_id, reason)); continue
            result = validate_new_board(board_id=board_id, filename=item["name"], split_payload=payload, source_records=source_records, posted_boards=posted_boards)
            if not result["ok"]:
                blocked.append((board_id, result["error"])); _append_unique_anomaly(new_anomalies, anomaly_seen, result["anomaly"]); continue
            for key, qty in result["deltas"].items(): delta_by_key[key] += qty
            for key, qty in result["board_sources"].items(): source_additions[key].append((board_id, qty))
            new_flows.extend(result["flows"]); new_board_records.append(result["board_record"]); _append_unique_anomaly(new_anomalies, anomaly_seen, result["anomaly"])
            accepted.append({**item, "board_id": board_id, "content_fingerprint": fp}); posted_boards.add(board_id); posted_keys.add((board_id, fp))
        timings["拆图结果核验"] = time.perf_counter()-t_verify

        t = time.perf_counter(); active = build_current_state(source_records, existing_state=existing["state"], delta_by_key=dict(delta_by_key), source_additions=dict(source_additions)); active.extend(frozen_rows); cumulative_rows = active + completed_history
        for row in active:
            if int(row.get("remaining", 0)) < 0: _append_unique_anomaly(new_anomalies, anomaly_seen, _overprocess_anomaly(row))
        timings["状态重建"] = time.perf_counter()-t

        all_flows = existing["flows"] + new_flows; all_boards = existing["board_records"] + new_board_records; all_anomalies = existing["anomalies"] + new_anomalies
        run_note = f"增量执行：订单源{len(source_files)}，新增{len(changes['added'])}，修改{len(changes['modified'])}，移除{len(changes['removed'])}；待处理板材{len(pending_files)}，新入账{len(accepted)}，补归档{len(reconciled)}，阻断{len(blocked)}。"
        cumulative_out, pending_out = temp_dir/CUMULATIVE_NAME, temp_dir/PENDING_NAME
        t = time.perf_counter(); generate_cumulative_report(cumulative_rows, all_flows, all_boards, all_anomalies, cumulative_out, note=run_note); generate_pending_report(active, pending_out, note=run_note); timings["Excel生成"] = time.perf_counter()-t
        t = time.perf_counter(); local_ledger = _validate_output(cumulative_out, pending_out, cumulative_rows, accepted); timings["本地验证"] = time.perf_counter()-t

        logger.info("拆图结果｜根目录待处理=%d｜新入账=%d｜重复补归档=%d｜阻断=%d", len(pending_files), len(accepted), len(reconciled), len(blocked))
        logger.info("本轮Drive正文下载文件总数=%d｜订单=%d｜累计台账=%d｜拆图结果=%d", source_downloads+ledger_downloads+pending_downloads, source_downloads, ledger_downloads, pending_downloads)

        if test_mode:
            logger.info("dry-run完成：不写正式Drive、不移动文件、不覆盖正式runtime cache"); return

        # 只有确有输出变化时上传；纯补归档且台账未变可直接依赖已确认的累计MD5。
        output_changed = bool(accepted or blocked or new_anomalies or changes['added'] or changes['modified'] or changes['removed'] or not cumulative_meta or not pending_meta)
        cumulative_verified = cumulative_meta
        pending_verified = pending_meta
        if output_changed:
            t = time.perf_counter(); cw = write_formal_file(drive, cumulative_meta, cumulative_out); timings["Drive上传"] += time.perf_counter()-t
            t = time.perf_counter(); cumulative_verified = drive.verify_remote_hash(cw["id"], cumulative_out); timings["metadata/hash验证"] += time.perf_counter()-t
            try:
                t = time.perf_counter(); pw = write_formal_file(drive, pending_meta, pending_out); timings["Drive上传"] += time.perf_counter()-t
                t = time.perf_counter(); pending_verified = drive.verify_remote_hash(pw["id"], pending_out); timings["metadata/hash验证"] += time.perf_counter()-t
            except Exception:
                logger.error("累计台账已确认上传，但当前待加工零件提交失败；不归档、不保存本轮runtime state，下轮将按累计台账MD5自动恢复")
                raise
        else:
            logger.info("正式Excel内容无变化，上传=0；使用已确认的累计台账MD5作为补归档依据")

        # 两份提交确认后才移动新板和补归档板。
        for item in accepted + reconciled:
            drive.move_file(item["id"], drive.archive_folder_id)
            state_text = "已累计、已录入" if item in accepted else "已累计、仅补归档"
            logger.info("板材处理结果｜文件=%s｜板材=%s｜状态=%s｜原因=%s｜处理建议=无", item["name"], item["board_id"], state_text,
                        "首次校验通过并已提交" if item in accepted else "累计台账已有相同内容记录，本次未重复累计")
        for board_id, reason in blocked:
            logger.warning("板材处理结果｜文件=%s_完成｜板材=%s｜状态=未累计、未记录｜原因=%s｜处理建议=核对基础资料后重新执行。", board_id, board_id, reason)

        # runtime cache 只在正式事务成功后更新。远端MD5已与本地一致，因此可直接保存本地解析状态。
        if cumulative_verified:
            state.set_ledger(cumulative_verified, local_ledger)
        state.save()

    timings["总耗时"] = time.perf_counter()-started
    logger.info("时间统计｜metadata扫描=%.3fs｜cache恢复=%.3fs｜Drive下载=%.3fs｜Excel解析=%.3fs｜拆图结果核验=%.3fs｜状态重建=%.3fs｜Excel生成=%.3fs｜本地验证=%.3fs｜Drive上传=%.3fs｜metadata/hash验证=%.3fs｜总耗时=%.3fs",
                timings['metadata扫描'], timings['cache恢复'], timings['Drive下载'], timings['Excel解析'], timings['拆图结果核验'], timings['状态重建'], timings['Excel生成'], timings['本地验证'], timings['Drive上传'], timings['metadata/hash验证'], timings['总耗时'])


if __name__ == "__main__": run()
