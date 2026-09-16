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
from copy import deepcopy
from pathlib import Path

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
    """为旧结果解析器补一条标准“板材 X 阻断”日志，且不重复输出。"""
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
    """测试/兼容辅助：只读 metadata 确认远端内容与本地文件字节一致。"""
    path = Path(local_path)
    metadata = drive.get_file(file_id, fields="id,name,modifiedTime,md5Checksum,size")
    remote_md5 = str(metadata.get("md5Checksum") or "").lower()
    local_md5 = _file_md5(path)
    if not remote_md5:
        raise RuntimeError(f"Drive 上传验证失败：{path.name} 未返回 md5Checksum")
    if remote_md5 != local_md5:
        raise RuntimeError(
            f"Drive 上传验证失败：{path.name} MD5 不一致，本地={local_md5}，远端={remote_md5}"
        )
    remote_size = metadata.get("size")
    if remote_size not in (None, "") and int(remote_size) != path.stat().st_size:
        raise RuntimeError(
            f"Drive 上传验证失败：{path.name} size 不一致，本地={path.stat().st_size}，远端={remote_size}"
        )
    return metadata


def _scalar(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        number = float(value)
        if math.isfinite(number):
            return round(number, 9)
        return str(value)
    return str(value).strip()


def _normalized_dict(row: dict, fields: tuple[str, ...] | None = None) -> dict:
    keys = fields if fields is not None else tuple(sorted(str(key) for key in row.keys()))
    return {key: _scalar(row.get(key)) for key in keys}


def _sorted_records(rows: list[dict], fields: tuple[str, ...] | None = None) -> list[dict]:
    normalized = [_normalized_dict(row, fields) for row in rows]
    return sorted(
        normalized,
        key=lambda item: json.dumps(item, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
    )


def business_ledger_signature(ledger: dict) -> str:
    """比较正式业务事实与正式状态；异常日志本身不触发正式表写回。"""
    payload = {
        "rows": _sorted_records(list(ledger.get("historical_rows", [])), _ROW_FIELDS),
        "flows": _sorted_records(list(ledger.get("flows", []))),
        "board_records": _sorted_records(list(ledger.get("board_records", []))),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def business_outputs_changed(existing: dict, candidate: dict) -> bool:
    return business_ledger_signature(existing) != business_ledger_signature(candidate)


def _scan_order_sources(drive) -> tuple[list[dict], list[dict]]:
    """只做一次完整订单 metadata 扫描，同时把订单目录数与有效源文件数分开记录。"""
    sources = drive.list_order_source_files()
    containers = drive.list_order_containers()  # 只多一次“正在加工”根目录 metadata，不重复扫25个订单正文目录。
    source_folder_ids = {str(item.get("order_folder_id", "")) for item in sources}
    missing = [item for item in containers if str(item.get("id", "")) not in source_folder_ids]
    logger.info(f"正在加工订单目录 {len(containers)} 个")
    logger.info(f"其中识别到订单原始汇总表 {len(sources)} 个")
    for item in missing:
        name = item.get("name", "")
        logger.warning(f"订单目录未找到正式汇总表，未纳入零件事实源：{name}")
        logger.warning(f"订单源未完成 {name}: 未找到正式汇总表")
    return sources, missing


def _all_sources_cached(source_files: list[dict], source_cache_path: Path) -> bool:
    if not source_files or not source_cache_path.exists():
        return False
    cache = SourceRecordCache(source_cache_path)
    return all(cache.get(item) is not None for item in source_files)


def run() -> None:
    from modules.drive_manager import DriveManager

    drive = DriveManager()
    if not drive.check_config():
        raise RuntimeError("Google Drive 配置不完整")

    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    runtime_path = Path(os.getenv("RUNTIME_STATE_FILE", "runtime_state/state.json"))
    source_cache_path = Path(os.getenv("SOURCE_RECORD_CACHE_FILE", ".cache/source_records.json"))
    runtime_state = RuntimeState(runtime_path)
    if runtime_state.restore_source_cache(source_cache_path):
        logger.info("runtime state 已恢复订单源解析缓存")

    # 只做一次昂贵的订单目录 metadata 扫描；后续 main.py 复用本轮快照。
    source_files, _missing_containers = _scan_order_sources(drive)
    pending_files = drive.list_pending_split_files()
    logger.info(f"拆图结果根目录待处理完成文件 {len(pending_files)} 个")

    official = prepare_official_files(drive)
    missing = [name for name, item in official.items() if item is None]
    if missing:
        if test_mode:
            return
        raise RuntimeError(f"正式台账自愈失败：{missing}")

    cumulative_meta = official[CUMULATIVE_NAME]
    pending_meta = official[PENDING_NAME]
    cached_ledger = runtime_state.get_cumulative_ledger(cumulative_meta)
    source_cache_hit = _all_sources_cached(source_files, source_cache_path)

    logger.info(
        "累计台账性能缓存："
        + ("命中，Drive MD5 与 runtime state 一致" if cached_ledger is not None else "未命中，将按事实源初始化")
    )

    # 真正无变化：订单解析缓存全部命中 + 累计MD5命中 + 根目录无板材 + 两正式表存在。
    # 此时结果视图不会发生任何业务变化，不生成/上传Excel，也不读取累计Excel正文。
    if cached_ledger is not None and source_cache_hit and not pending_files and not missing:
        logger.info(
            "无变化快速退出｜订单Excel正文下载=0｜累计Excel正文下载=0｜"
            "拆图完成文件正文下载=0｜正式Excel上传=0"
        )
        logger.info("本轮Drive正文下载文件总数 0（订单源0、累计台账0、拆图结果0）")
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

    seen_legacy_block_boards: set[str] = set()
    context: dict[str, object] = {
        "existing": None,
        "business_changed": None,
        "candidate_ledger": None,
        "cumulative_file_id": str(cumulative_meta["id"]),
        "ledger_cache_hit": cached_ledger is not None,
    }

    def capture_existing(path):
        path_obj = Path(path)
        if cached_ledger is not None and path_obj.name == "累计加工台账_当前.xlsx":
            ledger = deepcopy(cached_ledger)
            logger.info("累计加工台账：直接使用 runtime state 解析结果，未下载正式Excel正文")
        else:
            ledger = original_read_existing(path)
        if context["existing"] is None:
            context["existing"] = deepcopy(ledger)
        return ledger

    def guarded_download(self, file_id, save_path):
        target = Path(save_path)
        if (
            cached_ledger is not None
            and str(file_id) == str(cumulative_meta["id"])
            and target.name == "累计加工台账_当前.xlsx"
        ):
            # capture_existing 会直接返回 parsed ledger；这里故意不创建伪 Excel。
            return str(target)
        return original_download(self, file_id, save_path)

    def guarded_write_formal(drive_obj, current, local_path):
        name = Path(local_path).name
        if name == CUMULATIVE_NAME:
            candidate = original_read_existing(local_path)
            context["candidate_ledger"] = deepcopy(candidate)
            if current is None or context["existing"] is None:
                context["business_changed"] = True
                result = original_write_formal(drive_obj, current, local_path)
                context["cumulative_file_id"] = str(result["id"])
                return result
            changed = business_outputs_changed(context["existing"], candidate)
            context["business_changed"] = changed
            if not changed:
                logger.info("写回保护｜累计业务事实无变化：跳过累计加工台账上传")
                return {"id": current["id"], "_skipped": True}
            logger.info("写回保护｜检测到正式业务/状态变化：允许更新累计加工台账")
            result = original_write_formal(drive_obj, current, local_path)
            context["cumulative_file_id"] = str(result["id"])
            return result

        if name == PENDING_NAME:
            if current is not None and context["business_changed"] is False:
                logger.info("写回保护｜累计业务事实无变化：跳过当前待加工零件上传")
                return {"id": current["id"], "_skipped": True}
            return original_write_formal(drive_obj, current, local_path)
        return original_write_formal(drive_obj, current, local_path)

    def compatible_warning(message, *args, **kwargs):
        original_warning(message, *args, **kwargs)
        compat = legacy_block_compat_message(message, seen_legacy_block_boards)
        if compat:
            original_warning(compat)

    def compatible_info(message, *args, **kwargs):
        original_info(str(message), *args, **kwargs)

    # 将 safe_main 已完成的 metadata 扫描快照交给 main.py，避免再次遍历25个订单目录。
    business_main.read_existing_ledger = capture_existing
    business_main.write_formal_file = guarded_write_formal
    business_main.DriveManager.download_file = guarded_download
    business_main.DriveManager.list_order_source_files = lambda _self: deepcopy(source_files)
    business_main.DriveManager.list_pending_split_files = lambda _self: deepcopy(pending_files)
    business_main.discover_formal_files = lambda _drive: {
        "cumulative": deepcopy(cumulative_meta),
        "pending": deepcopy(pending_meta),
    }
    business_main.logger.warning = compatible_warning
    business_main.logger.info = compatible_info

    try:
        business_main.run()

        # runtime state 只在整个 production 事务成功返回后提交。
        # 若累计上传成功但待加工上传/归档失败，这里不会执行；下一轮会因MD5不一致回读正式累计台账恢复。
        if not test_mode:
            candidate_ledger = context.get("candidate_ledger")
            if isinstance(candidate_ledger, dict):
                cumulative_id = str(context.get("cumulative_file_id") or cumulative_meta["id"])
                final_meta = drive.get_file(
                    cumulative_id,
                    fields="id,name,modifiedTime,md5Checksum,size",
                )
                runtime_state.set_cumulative_ledger(final_meta, candidate_ledger)
                runtime_state.capture_source_cache(source_cache_path)
                runtime_state.save()
                logger.info(f"runtime state 已提交：{runtime_path}")
    finally:
        business_main.DriveManager.download_file = original_download
        business_main.DriveManager.list_order_source_files = original_list_sources
        business_main.DriveManager.list_pending_split_files = original_list_pending
        business_main.discover_formal_files = original_discover_formal
        business_main.logger.warning = original_warning
        business_main.logger.info = original_info


if __name__ == "__main__":
    run()
