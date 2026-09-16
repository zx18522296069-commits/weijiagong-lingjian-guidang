"""未加工零件安全入口。

职责：
- 只做正式写回保护、runtime state 性能缓存、快速读取和旧结果日志兼容；
- Google Drive 正式订单与累计台账始终是事实源；
- runtime state 缺失/损坏/MD5不一致时自动退回真实 Drive 初始化；
- 无变化且无待处理板材时，在 metadata 扫描后直接退出，不下载正式 Excel 正文。
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import time
from collections import defaultdict
from copy import deepcopy
from pathlib import Path

from modules.excel_generator import REPORT_SCHEMA_VERSION
from modules.fast_ledger import read_existing_ledger_fast
from modules.formal_files import CUMULATIVE_NAME, PENDING_NAME
from modules.logger import get_logger
from modules.runtime_state import RuntimeState
from modules.source_cache import SourceRecordCache
from self_healing_main import prepare_official_files

logger = get_logger()

_ROW_FIELDS = (
    "order", "drawing", "thickness", "bevel", "length", "width", "quantity",
    "total_weight_t", "board_sources", "processed", "remaining", "pending_weight_t", "status",
)
_LEGACY_BLOCK_RE = re.compile(r"^板材\s+(.+?)\s+阻断：")
_STRUCTURED_BLOCK_RE = re.compile(
    r"^板材处理结果｜文件=[^｜]+｜板材=([^｜]+)｜状态=未累计、未记录｜原因=(.+?)｜处理建议="
)


def legacy_block_compat_message(message: object, seen_boards: set[str]) -> str | None:
    text = str(message)
    legacy = _LEGACY_BLOCK_RE.match(text)
    if legacy:
        seen_boards.add(legacy.group(1).strip())
        return None
    structured = _STRUCTURED_BLOCK_RE.match(text)
    if not structured:
        return None
    board_id = structured.group(1).strip()
    reason = structured.group(2).strip()
    if not board_id or board_id in seen_boards:
        return None
    seen_boards.add(board_id)
    return f"板材 {board_id} 阻断：{reason}"


def _file_md5(path: str | Path) -> str:
    digest = hashlib.md5()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def verify_uploaded_file_metadata(drive, file_id: str, local_path: str | Path) -> dict:
    path = Path(local_path)
    metadata = drive.get_file(file_id, fields="id,name,modifiedTime,md5Checksum,size")
    remote_md5 = str(metadata.get("md5Checksum") or "").lower()
    local_md5 = _file_md5(path)
    if not remote_md5:
        raise RuntimeError(f"Drive 上传验证失败：{path.name} 未返回 md5Checksum")
    if remote_md5 != local_md5:
        raise RuntimeError(f"Drive 上传验证失败：{path.name} MD5 不一致，本地={local_md5}，远端={remote_md5}")
    remote_size = metadata.get("size")
    if remote_size not in (None, "") and int(remote_size) != path.stat().st_size:
        raise RuntimeError(f"Drive 上传验证失败：{path.name} size 不一致，本地={path.stat().st_size}，远端={remote_size}")
    return metadata


def _scalar(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        number = float(value)
        return round(number, 9) if math.isfinite(number) else str(value)
    return str(value).strip()


def _normalized_dict(row: dict, fields: tuple[str, ...] | None = None) -> dict:
    keys = fields if fields is not None else tuple(sorted(str(key) for key in row.keys()))
    return {key: _scalar(row.get(key)) for key in keys}


def _sorted_records(rows: list[dict], fields: tuple[str, ...] | None = None) -> list[dict]:
    normalized = [_normalized_dict(row, fields) for row in rows]
    return sorted(normalized, key=lambda item: json.dumps(item, ensure_ascii=False, sort_keys=True, separators=(",", ":")))


def business_ledger_signature(ledger: dict) -> str:
    """正式业务事实+报表结构签名；普通阻断日志不触发写回，超加工异常属于永久业务事实。"""
    business_anomalies = [
        row for row in ledger.get("anomalies", [])
        if str(row.get("异常类型", "")).strip() == "超加工/待核查"
    ]
    payload = {
        "report_schema_version": int(ledger.get("report_schema_version", 1) or 1),
        "rows": _sorted_records(list(ledger.get("historical_rows", [])), _ROW_FIELDS),
        "flows": _sorted_records(list(ledger.get("flows", []))),
        "board_records": _sorted_records(list(ledger.get("board_records", []))),
        "business_anomalies": _sorted_records(business_anomalies),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def business_outputs_changed(existing: dict, candidate: dict) -> bool:
    return business_ledger_signature(existing) != business_ledger_signature(candidate)


def _scan_order_sources(drive) -> tuple[list[dict], list[dict]]:
    sources = drive.list_order_source_files()
    containers = drive.list_order_containers()
    source_folder_ids = {str(item.get("order_folder_id", "")) for item in sources}
    missing = [item for item in containers if str(item.get("id", "")) not in source_folder_ids]
    logger.info(f"正在加工订单目录 {len(containers)} 个")
    logger.info(f"其中识别到订单原始汇总表 {len(sources)} 个")
    for item in missing:
        name = item.get("name", "")
        logger.warning(f"订单目录未找到正式汇总表，未纳入零件事实源：{name}")
        logger.warning(f"订单源未完成 {name}: 未找到正式汇总表")
    return sources, missing


def _source_change_counts(source_files: list[dict], source_cache_path: Path) -> dict[str, int]:
    cache = SourceRecordCache(source_cache_path)
    old_ids = set(cache.entries)
    current_ids = {str(item.get("id", "")) for item in source_files if item.get("id")}
    counts = {"unchanged": 0, "added": 0, "modified": 0, "removed": len(old_ids - current_ids)}
    for item in source_files:
        file_id = str(item.get("id", ""))
        entry = cache.entries.get(file_id)
        if entry is None:
            counts["added"] += 1
        elif entry.get("signature") == cache.signature(item):
            counts["unchanged"] += 1
        else:
            counts["modified"] += 1
    return counts


def _all_sources_cached(source_files: list[dict], source_cache_path: Path) -> bool:
    if not source_files or not source_cache_path.exists():
        return False
    cache = SourceRecordCache(source_cache_path)
    return all(cache.get(item) is not None for item in source_files)


def _download_category(path: Path) -> str:
    name = path.name
    if name.startswith("source_") or name.startswith("heal_source_"):
        return "订单源"
    if name == "累计加工台账_当前.xlsx":
        return "累计台账"
    if name.startswith("pending_"):
        return "拆图结果"
    return "其他"


def _log_performance(*, total_started: float, cache_seconds: float, metadata_seconds: float, timings: dict, downloads: dict, uploads: int) -> None:
    total = time.perf_counter() - total_started
    logger.info(
        "性能统计｜"
        f"cache读取={cache_seconds:.3f}s｜metadata扫描={metadata_seconds:.3f}s｜"
        f"Drive正文读取={timings.get('drive_download', 0.0):.3f}s｜Excel解析={timings.get('excel_parse', 0.0):.3f}s｜"
        f"拆图核验={timings.get('board_validate', 0.0):.3f}s｜状态重建={timings.get('state_rebuild', 0.0):.3f}s｜"
        f"Excel生成={timings.get('excel_generate', 0.0):.3f}s｜本地验证={timings.get('local_validate', 0.0):.3f}s｜"
        f"Drive上传={timings.get('drive_upload', 0.0):.3f}s｜metadata/hash验证={timings.get('metadata_hash', 0.0):.3f}s｜"
        f"总耗时={total:.3f}s"
    )
    total_downloads = sum(downloads.values())
    logger.info(
        f"本轮Drive正文下载文件总数 {total_downloads}（订单源{downloads.get('订单源', 0)}、"
        f"累计台账{downloads.get('累计台账', 0)}、拆图结果{downloads.get('拆图结果', 0)}、其他{downloads.get('其他', 0)}）"
    )
    logger.info(f"本轮正式Excel上传次数 {uploads}")


def run() -> None:
    from modules.drive_manager import DriveManager

    total_started = time.perf_counter()
    timings = defaultdict(float)
    downloads = defaultdict(int)
    counters = {"uploads": 0}

    drive = DriveManager()
    if not drive.check_config():
        raise RuntimeError("Google Drive 配置不完整")

    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    runtime_path = Path(os.getenv("RUNTIME_STATE_FILE", "runtime_state/state.json"))
    source_cache_path = Path(os.getenv("SOURCE_RECORD_CACHE_FILE", ".cache/source_records.json"))

    cache_started = time.perf_counter()
    runtime_state = RuntimeState(runtime_path)
    if runtime_state.restore_source_cache(source_cache_path):
        logger.info("runtime state 已恢复订单源解析缓存")
    cache_seconds = time.perf_counter() - cache_started

    metadata_started = time.perf_counter()
    source_files, _missing_containers = _scan_order_sources(drive)
    pending_files = drive.list_pending_split_files()
    logger.info(f"拆图结果根目录待处理完成文件 {len(pending_files)} 个")
    official = prepare_official_files(drive)
    metadata_seconds = time.perf_counter() - metadata_started

    source_changes = _source_change_counts(source_files, source_cache_path)
    logger.info(
        "Drive扫描｜"
        f"当前订单源={len(source_files)}｜unchanged={source_changes['unchanged']}｜added={source_changes['added']}｜"
        f"modified={source_changes['modified']}｜removed={source_changes['removed']}"
    )

    missing = [name for name, item in official.items() if item is None]
    if missing:
        if test_mode:
            _log_performance(total_started=total_started, cache_seconds=cache_seconds, metadata_seconds=metadata_seconds, timings=timings, downloads=downloads, uploads=0)
            return
        raise RuntimeError(f"正式台账自愈失败：{missing}")

    cumulative_meta = official[CUMULATIVE_NAME]
    pending_meta = official[PENDING_NAME]
    cached_ledger = runtime_state.get_cumulative_ledger(cumulative_meta)
    source_cache_hit = _all_sources_cached(source_files, source_cache_path)
    cached_schema_version = int((cached_ledger or {}).get("report_schema_version", 1) or 1)
    schema_outdated = cached_ledger is not None and cached_schema_version < REPORT_SCHEMA_VERSION
    logger.info(
        "累计台账｜cache=" + ("命中" if cached_ledger is not None else "未命中")
        + f"｜Drive MD5变化={'否' if cached_ledger is not None else '是/未知'}｜是否下载正式累计Excel={'否' if cached_ledger is not None else '需要初始化'}"
        + f"｜报表格式V{cached_schema_version}->V{REPORT_SCHEMA_VERSION}{'（需升级）' if schema_outdated else ''}"
    )

    if schema_outdated:
        logger.info("报表格式升级：即使业务数量无变化，也将重新生成并刷新两份正式 Excel")

    if cached_ledger is not None and source_cache_hit and not pending_files and not missing and not schema_outdated:
        logger.info("无变化快速退出｜订单Excel正文下载=0｜累计Excel正文下载=0｜拆图完成文件正文下载=0｜正式Excel上传=0")
        _log_performance(total_started=total_started, cache_seconds=cache_seconds, metadata_seconds=metadata_seconds, timings=timings, downloads=downloads, uploads=0)
        return

    os.environ["CUMULATIVE_FILE_ID"] = str(cumulative_meta["id"])
    os.environ["PENDING_FILE_ID"] = str(pending_meta["id"])

    import main as business_main

    original_read_existing = read_existing_ledger_fast
    original_write_formal = business_main.write_formal_file
    original_download = business_main.DriveManager.download_file
    original_list_sources = business_main.DriveManager.list_order_source_files
    original_list_pending = business_main.DriveManager.list_pending_split_files
    original_discover_formal = business_main.discover_formal_files
    original_warning = business_main.logger.warning
    original_info = business_main.logger.info
    original_read_source = business_main.read_source_summary
    original_read_split = business_main.read_split_result
    original_validate = business_main.validate_new_board
    original_reconcile = business_main.reconcile_posted_board
    original_build_state = business_main.build_current_state
    original_gen_cumulative = business_main.generate_cumulative_report
    original_gen_pending = business_main.generate_pending_report
    original_verify_local = business_main._verify_output_workbooks
    original_verify_upload = business_main._verify_drive_upload

    seen_legacy_block_boards: set[str] = set()
    context: dict[str, object] = {
        "existing": None,
        "business_changed": None,
        "candidate_ledger": None,
        "cumulative_file_id": str(cumulative_meta["id"]),
    }

    def _timed(bucket, fn, *args, **kwargs):
        started = time.perf_counter()
        try:
            return fn(*args, **kwargs)
        finally:
            timings[bucket] += time.perf_counter() - started

    def capture_existing(path):
        path_obj = Path(path)
        if cached_ledger is not None and path_obj.name == "累计加工台账_当前.xlsx":
            ledger = deepcopy(cached_ledger)
            logger.info("累计加工台账：直接使用 runtime state 解析结果，未下载正式Excel正文")
        else:
            ledger = _timed("excel_parse", original_read_existing, path)
        if context["existing"] is None:
            context["existing"] = deepcopy(ledger)
        return ledger

    def guarded_download(self, file_id, save_path):
        target = Path(save_path)
        if cached_ledger is not None and str(file_id) == str(cumulative_meta["id"]) and target.name == "累计加工台账_当前.xlsx":
            return str(target)
        category = _download_category(target)
        downloads[category] += 1
        return _timed("drive_download", original_download, self, file_id, save_path)

    def _write(drive_obj, current, local_path):
        counters["uploads"] += 1
        return _timed("drive_upload", original_write_formal, drive_obj, current, local_path)

    def guarded_write_formal(drive_obj, current, local_path):
        name = Path(local_path).name
        if name == CUMULATIVE_NAME:
            candidate = _timed("excel_parse", original_read_existing, local_path)
            context["candidate_ledger"] = deepcopy(candidate)
            if current is None or context["existing"] is None:
                context["business_changed"] = True
                result = _write(drive_obj, current, local_path)
                context["cumulative_file_id"] = str(result["id"])
                return result
            changed = business_outputs_changed(context["existing"], candidate)
            context["business_changed"] = changed
            if not changed:
                logger.info("写回保护｜累计业务事实与报表格式均无变化：跳过累计加工台账上传")
                return {"id": current["id"], "_skipped": True}
            logger.info("写回保护｜检测到正式业务/状态/报表格式变化：允许更新累计加工台账")
            result = _write(drive_obj, current, local_path)
            context["cumulative_file_id"] = str(result["id"])
            return result
        if name == PENDING_NAME:
            if current is not None and context["business_changed"] is False:
                logger.info("写回保护｜累计业务事实与报表格式均无变化：跳过当前待加工零件上传")
                return {"id": current["id"], "_skipped": True}
            return _write(drive_obj, current, local_path)
        return _write(drive_obj, current, local_path)

    def compatible_warning(message, *args, **kwargs):
        original_warning(message, *args, **kwargs)
        compat = legacy_block_compat_message(message, seen_legacy_block_boards)
        if compat:
            original_warning(compat)

    business_main.read_existing_ledger = capture_existing
    business_main.write_formal_file = guarded_write_formal
    business_main.DriveManager.download_file = guarded_download
    business_main.DriveManager.list_order_source_files = lambda _self: deepcopy(source_files)
    business_main.DriveManager.list_pending_split_files = lambda _self: deepcopy(pending_files)
    business_main.discover_formal_files = lambda _drive: {"cumulative": deepcopy(cumulative_meta), "pending": deepcopy(pending_meta)}
    business_main.read_source_summary = lambda *a, **k: _timed("excel_parse", original_read_source, *a, **k)
    business_main.read_split_result = lambda *a, **k: _timed("excel_parse", original_read_split, *a, **k)
    business_main.validate_new_board = lambda *a, **k: _timed("board_validate", original_validate, *a, **k)
    business_main.reconcile_posted_board = lambda *a, **k: _timed("board_validate", original_reconcile, *a, **k)
    business_main.build_current_state = lambda *a, **k: _timed("state_rebuild", original_build_state, *a, **k)
    business_main.generate_cumulative_report = lambda *a, **k: _timed("excel_generate", original_gen_cumulative, *a, **k)
    business_main.generate_pending_report = lambda *a, **k: _timed("excel_generate", original_gen_pending, *a, **k)
    business_main._verify_output_workbooks = lambda *a, **k: _timed("local_validate", original_verify_local, *a, **k)
    business_main._verify_drive_upload = lambda *a, **k: _timed("metadata_hash", original_verify_upload, *a, **k)
    business_main.logger.warning = compatible_warning
    business_main.logger.info = lambda message, *args, **kwargs: original_info(str(message), *args, **kwargs)

    try:
        business_main.run()
        if not test_mode:
            candidate_ledger = context.get("candidate_ledger")
            if isinstance(candidate_ledger, dict):
                cumulative_id = str(context.get("cumulative_file_id") or cumulative_meta["id"])
                final_meta = drive.get_file(cumulative_id, fields="id,name,modifiedTime,md5Checksum,size")
                runtime_state.set_cumulative_ledger(final_meta, candidate_ledger)
                runtime_state.capture_source_cache(source_cache_path)
                runtime_state.save()
                logger.info(f"runtime state 已提交：{runtime_path}")
        _log_performance(total_started=total_started, cache_seconds=cache_seconds, metadata_seconds=metadata_seconds, timings=timings, downloads=downloads, uploads=counters["uploads"])
    finally:
        business_main.DriveManager.download_file = original_download
        business_main.DriveManager.list_order_source_files = original_list_sources
        business_main.DriveManager.list_pending_split_files = original_list_pending
        business_main.discover_formal_files = original_discover_formal
        business_main.read_source_summary = original_read_source
        business_main.read_split_result = original_read_split
        business_main.validate_new_board = original_validate
        business_main.reconcile_posted_board = original_reconcile
        business_main.build_current_state = original_build_state
        business_main.generate_cumulative_report = original_gen_cumulative
        business_main.generate_pending_report = original_gen_pending
        business_main._verify_output_workbooks = original_verify_local
        business_main._verify_drive_upload = original_verify_upload
        business_main.logger.warning = original_warning
        business_main.logger.info = original_info


if __name__ == "__main__":
    run()
