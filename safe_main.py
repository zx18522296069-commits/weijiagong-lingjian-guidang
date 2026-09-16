"""未加工零件安全入口。

在现有稳定业务主流程外增加一层正式写回保护与兼容层：
- 阻断板只记录运行日志，不因为新增阻断异常/运行说明而改写正式 Excel；
- 只有订单事实、累计数量、加工流水、板材入账记录或正式状态规范化真正变化时，才允许覆盖两份正式表；
- 补归档场景可在不重写正式表的情况下继续完成归档；
- 累计台账采用只读顺序扫描，避免 openpyxl read_only 随机访问反复扫 XML；
- 正式 Excel 上传后只做 Drive metadata / MD5 / size 确认，不再重新下载整个 Excel；
- 不改变 main.py 现有的确定性匹配、整板校验、幂等和归档规则。
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import shutil
from copy import deepcopy
from pathlib import Path

from modules.fast_ledger import read_existing_ledger_fast
from modules.formal_files import CUMULATIVE_NAME, PENDING_NAME
from modules.logger import get_logger
from self_healing_main import prepare_official_files

logger = get_logger()


_ROW_FIELDS = (
    "order",
    "drawing",
    "thickness",
    "bevel",
    "length",
    "width",
    "quantity",
    "total_weight_t",
    "board_sources",
    "processed",
    "remaining",
    "pending_weight_t",
    "status",
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
    """只读取 Drive metadata，确认远端内容与本地上传文件字节一致。"""
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
    """比较正式业务事实与正式状态；故意忽略“异常记录”和运行说明。"""
    payload = {
        "rows": _sorted_records(list(ledger.get("historical_rows", [])), _ROW_FIELDS),
        "flows": _sorted_records(list(ledger.get("flows", []))),
        "board_records": _sorted_records(list(ledger.get("board_records", []))),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def business_outputs_changed(existing: dict, candidate: dict) -> bool:
    return business_ledger_signature(existing) != business_ledger_signature(candidate)


def _log_order_source_coverage(drive) -> None:
    """明确区分订单目录总数与真正找到的原始汇总表数，缺失项不得静默消失。"""
    containers = drive.list_order_containers()
    sources = drive.list_order_source_files()
    source_folder_ids = {str(item.get("order_folder_id", "")) for item in sources}
    missing = [item for item in containers if str(item.get("id", "")) not in source_folder_ids]

    logger.info(f"正在加工订单目录 {len(containers)} 个")
    logger.info(f"其中识别到订单原始汇总表 {len(sources)} 个")
    for item in missing:
        name = item.get("name", "")
        logger.warning(f"订单目录未找到正式汇总表，未纳入零件事实源：{name}")
        # 兼容控制台现有结果解析器；这是订单源异常，不得计入板材数量。
        logger.warning(f"订单源未完成 {name}: 未找到正式汇总表")


def run() -> None:
    from modules.drive_manager import DriveManager

    drive = DriveManager()
    if not drive.check_config():
        raise RuntimeError("Google Drive 配置不完整")

    _log_order_source_coverage(drive)

    official = prepare_official_files(drive)
    missing = [name for name, item in official.items() if item is None]
    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    if missing:
        if test_mode:
            return
        raise RuntimeError(f"正式台账自愈失败：{missing}")

    os.environ["CUMULATIVE_FILE_ID"] = str(official[CUMULATIVE_NAME]["id"])
    os.environ["PENDING_FILE_ID"] = str(official[PENDING_NAME]["id"])

    import main as business_main

    business_main.read_existing_ledger = read_existing_ledger_fast
    original_read_existing = read_existing_ledger_fast
    original_write_formal = business_main.write_formal_file
    original_download = business_main.DriveManager.download_file
    original_warning = business_main.logger.warning
    original_info = business_main.logger.info
    seen_legacy_block_boards: set[str] = set()
    verification_sources: dict[str, Path] = {}
    context: dict[str, object] = {
        "existing": None,
        "business_changed": None,
    }

    def capture_existing(path):
        ledger = original_read_existing(path)
        if context["existing"] is None:
            context["existing"] = deepcopy(ledger)
        return ledger

    def _write_and_verify(drive_obj, current, local_path):
        result = original_write_formal(drive_obj, current, local_path)
        file_id = str(result["id"])
        verify_uploaded_file_metadata(drive_obj, file_id, local_path)
        verification_sources[file_id] = Path(local_path)
        original_info(f"上传校验通过｜{Path(local_path).name}｜Drive metadata/MD5/size 一致")
        return result

    def guarded_write_formal(drive_obj, current, local_path):
        name = Path(local_path).name

        if name == CUMULATIVE_NAME:
            if current is None or context["existing"] is None:
                context["business_changed"] = True
                return _write_and_verify(drive_obj, current, local_path)

            candidate = original_read_existing(local_path)
            changed = business_outputs_changed(context["existing"], candidate)
            context["business_changed"] = changed
            if not changed:
                logger.info(
                    "写回保护｜累计业务事实无变化：跳过累计加工台账上传；"
                    "阻断异常仅保留运行日志，不改正式台账"
                )
                verification_sources[str(current["id"])] = Path(local_path)
                return {"id": current["id"], "_skipped": True}
            logger.info("写回保护｜检测到正式业务/状态变化：允许更新累计加工台账")
            return _write_and_verify(drive_obj, current, local_path)

        if name == PENDING_NAME:
            if current is not None and context["business_changed"] is False:
                logger.info("写回保护｜累计业务事实无变化：跳过当前待加工零件上传")
                verification_sources[str(current["id"])] = Path(local_path)
                return {"id": current["id"], "_skipped": True}
            return _write_and_verify(drive_obj, current, local_path)

        return original_write_formal(drive_obj, current, local_path)

    def no_redownload_verification(self, file_id, save_path):
        source = verification_sources.get(str(file_id))
        target = Path(save_path)
        if source is not None and target.name.startswith("verify_"):
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)
            return str(target)
        return original_download(self, file_id, save_path)

    def compatible_warning(message, *args, **kwargs):
        original_warning(message, *args, **kwargs)
        compat = legacy_block_compat_message(message, seen_legacy_block_boards)
        if compat:
            original_warning(compat)

    def compatible_info(message, *args, **kwargs):
        text = str(message)
        if "编号重复但内容不同：按新板材继续校验并分别入账" in text:
            text = text.replace(
                "编号重复但内容不同：按新板材继续校验并分别入账",
                "编号重复且内容不同：进入同板材号内容冲突校验，禁止自动再次入账",
            )
        if text == "正式文件回读验证通过":
            text = "正式文件验证通过：本地语义校验完成，上传文件已通过 Drive metadata/MD5/size 确认；未重新下载整个 Excel"
        original_info(text, *args, **kwargs)

    business_main.read_existing_ledger = capture_existing
    business_main.write_formal_file = guarded_write_formal
    business_main.DriveManager.download_file = no_redownload_verification
    business_main.logger.warning = compatible_warning
    business_main.logger.info = compatible_info
    try:
        business_main.run()
    finally:
        business_main.DriveManager.download_file = original_download
        business_main.logger.warning = original_warning
        business_main.logger.info = original_info


if __name__ == "__main__":
    run()
