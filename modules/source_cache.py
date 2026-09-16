"""订单原始汇总表的结构化解析缓存。

缓存只用于避免对未变化的订单源重复下载/解析；正式事实源仍然是 Google Drive
“正在加工”中的原始汇总表。当前缓存签名包含 file id、modifiedTime、md5Checksum、
父目录、文件名和订单目录名；任一变化都会重新下载并解析。

从 v1 升级到 v2 时允许一次安全迁移：只有旧签名中的 modifiedTime、size、文件名、
订单目录名仍全部一致，才复用旧解析结果，并立即把该条目升级为当前完整签名。
"""

from __future__ import annotations

import json
from pathlib import Path


CACHE_VERSION = 2
LEGACY_CACHE_VERSION = 1


class SourceRecordCache:
    def __init__(self, path: Path):
        self.path = path
        self.entries: dict[str, dict] = {}
        self.loaded_version: int | None = None
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return
        version = payload.get("version")
        if version not in {CACHE_VERSION, LEGACY_CACHE_VERSION}:
            return
        entries = payload.get("entries")
        if isinstance(entries, dict):
            self.entries = entries
            self.loaded_version = int(version)

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

    @staticmethod
    def _legacy_signature_matches(entry_signature: dict, item: dict) -> bool:
        if not isinstance(entry_signature, dict):
            return False
        expected = {
            "modifiedTime": str(item.get("modifiedTime", "")),
            "size": str(item.get("size", "")),
            "name": str(item.get("name", "")),
            "order_container_name": str(item.get("order_container_name", "")),
        }
        return all(str(entry_signature.get(key, "")) == value for key, value in expected.items())

    def get(self, item: dict) -> dict | None:
        file_id = str(item.get("id", ""))
        if not file_id:
            return None
        entry = self.entries.get(file_id)
        if not isinstance(entry, dict):
            return None

        current_signature = self.signature(item)
        entry_signature = entry.get("signature")
        if entry_signature != current_signature:
            if self.loaded_version == LEGACY_CACHE_VERSION and self._legacy_signature_matches(entry_signature, item):
                # 旧缓存只在旧签名完全一致时允许复用一次；随后保存即升级到 v2。
                entry["signature"] = current_signature
            else:
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
        self.loaded_version = CACHE_VERSION
