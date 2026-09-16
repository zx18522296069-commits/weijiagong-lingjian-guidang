"""“正在加工”根目录正式台账文件发现与创建辅助。"""

from __future__ import annotations

from pathlib import Path

from modules.drive_manager import FOLDER_MIME, SHORTCUT_MIME

CUMULATIVE_NAME = "累计加工台账.xlsx"
PENDING_NAME = "当前待加工零件.xlsx"


def find_unique_working_file(drive, filename: str) -> dict | None:
    """按正式文件名在“正在加工”根目录唯一查找；文件夹/快捷方式不能冒充正式表。"""
    matches = [
        item
        for item in drive.list_children(drive.working_folder_id)
        if str(item.get("name", "")).strip() == filename
        and item.get("mimeType") not in {FOLDER_MIME, SHORTCUT_MIME}
    ]
    if len(matches) > 1:
        ids = [item.get("id", "") for item in matches]
        raise RuntimeError(f"正在加工目录存在多个同名正式文件，禁止自动猜测: {filename} -> {ids}")
    return matches[0] if matches else None


def discover_formal_files(drive) -> dict[str, dict | None]:
    return {
        "cumulative": find_unique_working_file(drive, CUMULATIVE_NAME),
        "pending": find_unique_working_file(drive, PENDING_NAME),
    }


def empty_ledger() -> dict:
    """累计台账不存在时的空永久台账状态。"""
    return {
        "state": {},
        "historical_rows": [],
        "flows": [],
        "board_records": [],
        "anomalies": [],
        "posted_boards": set(),
        "posted_board_keys": set(),
        "legacy_posted_boards": set(),
    }


def write_formal_file(drive, current: dict | None, local_path: Path) -> dict:
    """有正式文件则覆盖；缺失则在“正在加工”根目录新建，并返回当前真实元数据。"""
    if current:
        return drive.update_file_content(current["id"], str(local_path))
    return drive.upload_file(str(local_path), drive.working_folder_id)
