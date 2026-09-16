"""订单原始汇总表的结构化解析缓存。

缓存只用于避免对未变化的订单源重复下载/解析；正式事实源仍然是 Google Drive
“正在加工”中的原始汇总表。缓存命中签名包含 file id、modifiedTime、md5Checksum、
父目录、文件名和订单目录名；任一变化都会重新下载并解析。
"""

from __future__ import annotations

import json
from pathlib import Path


CACHE_VERSION = 2


class SourceRecordCache:
    def __init__(self, path: Path):
        self.path = path
        self.entries: dict[str, dict] = {}
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return
        if payload.get("version") != CACHE_VERSION:
            return
        entries = payload.get("entries")
        if isinstance(entries, dict):
            self.entries = entries

    @staticmethod
    def signature(item: dict) -> dict:
        return {
            "file_id": str(item.get("id", "")),
            "modifiedTime": str(item.get("modifiedTime", "")),
            "md5Checksum": str(item.get("md5Checksum", "")),
            "parent_folder_id": str(item.get("parent_folder_id", "")),
            "name": str(item.get("name", "")),
            "order_container_name": str(item.get("order_container_name", "")),
        }

    def get(self, item: dict) -> dict | None:
        file_id = str(item.get("id", ""))
        if not file_id:
            return None
        entry = self.entries.get(file_id)
        if not isinstance(entry, dict):
            return None
        if entry.get("signature") != self.signature(item):
            return None
        status = entry.get("status")
        if status == "ok" and isinstance(entry.get("records"), list):
            return entry
        if status == "error" and isinstance(entry.get("error"), dict):
            return entry
        return None

    def store_records(self, item: dict, records: list[dict]) -> None:
        file_id = str(item["id"])
        self.entries[file_id] = {
            "signature": self.signature(item),
            "status": "ok",
            "records": records,
        }

    def store_error(self, item: dict, error: dict) -> None:
        file_id = str(item["id"])
        self.entries[file_id] = {
            "signature": self.signature(item),
            "status": "error",
            "error": error,
        }

    def retain(self, current_items: list[dict]) -> None:
        current_ids = {str(item.get("id", "")) for item in current_items if item.get("id")}
        self.entries = {
            file_id: entry
            for file_id, entry in self.entries.items()
            if file_id in current_ids
        }

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = self.path.with_suffix(self.path.suffix + ".tmp")
        temp_path.write_text(
            json.dumps(
                {"version": CACHE_VERSION, "entries": self.entries},
                ensure_ascii=False,
                separators=(",", ":"),
            ),
            encoding="utf-8",
        )
        temp_path.replace(self.path)
