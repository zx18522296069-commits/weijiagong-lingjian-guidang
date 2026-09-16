"""可重建的运行时性能缓存；正式事实仍来自 Google Drive。"""
from __future__ import annotations

import json
from pathlib import Path

STATE_SCHEMA_VERSION = 2


def _key_to_text(key: tuple) -> str:
    return json.dumps(list(key), ensure_ascii=False, separators=(",", ":"))


def _text_to_key(text: str) -> tuple:
    value = json.loads(text)
    return (str(value[0]), str(value[1]), float(value[2]), str(value[3]))


def serialize_ledger(ledger: dict) -> dict:
    return {
        "state": {_key_to_text(k): v for k, v in ledger.get("state", {}).items()},
        "historical_rows": ledger.get("historical_rows", []),
        "flows": ledger.get("flows", []),
        "board_records": ledger.get("board_records", []),
        "anomalies": ledger.get("anomalies", []),
        "posted_boards": sorted(ledger.get("posted_boards", set())),
        "posted_board_keys": [list(v) for v in sorted(ledger.get("posted_board_keys", set()))],
        "legacy_posted_boards": sorted(ledger.get("legacy_posted_boards", set())),
    }


def deserialize_ledger(data: dict) -> dict:
    return {
        "state": {_text_to_key(k): v for k, v in (data.get("state") or {}).items()},
        "historical_rows": list(data.get("historical_rows") or []),
        "flows": list(data.get("flows") or []),
        "board_records": list(data.get("board_records") or []),
        "anomalies": list(data.get("anomalies") or []),
        "posted_boards": set(data.get("posted_boards") or []),
        "posted_board_keys": {tuple(v) for v in (data.get("posted_board_keys") or [])},
        "legacy_posted_boards": set(data.get("legacy_posted_boards") or []),
    }


def source_signature(item: dict) -> dict:
    return {
        "file_id": str(item.get("id", "")),
        "name": str(item.get("name", "")),
        "modifiedTime": str(item.get("modifiedTime", "")),
        "md5Checksum": str(item.get("md5Checksum", "")),
        "parent_folder_id": str(item.get("parent_folder_id") or item.get("order_folder_id") or ""),
        "order_container_name": str(item.get("order_container_name", "")),
    }


class RuntimeState:
    def __init__(self, path: Path):
        self.path = path
        self.valid = False
        self.sources: dict[str, dict] = {}
        self.ledger: dict | None = None
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
            if payload.get("state_schema_version") != STATE_SCHEMA_VERSION:
                return
            if not isinstance(payload.get("sources"), dict):
                return
            self.sources = payload["sources"]
            self.ledger = payload.get("ledger") if isinstance(payload.get("ledger"), dict) else None
            self.valid = True
        except (OSError, ValueError, TypeError, KeyError):
            self.valid = False
            self.sources = {}
            self.ledger = None

    def classify_sources(self, current_items: list[dict]) -> dict:
        current = {str(item["id"]): item for item in current_items}
        unchanged, added, modified, removed = [], [], [], []
        for file_id, item in current.items():
            old = self.sources.get(file_id)
            if old is None:
                added.append(item)
            elif old.get("signature") == source_signature(item):
                unchanged.append(item)
            else:
                modified.append(item)
        for file_id, old in self.sources.items():
            if file_id not in current:
                removed.append(old)
        return {"unchanged": unchanged, "added": added, "modified": modified, "removed": removed}

    def get_source(self, item: dict) -> dict | None:
        entry = self.sources.get(str(item.get("id", "")))
        if not entry or entry.get("signature") != source_signature(item):
            return None
        return entry

    def set_source_records(self, item: dict, records: list[dict]) -> None:
        self.sources[str(item["id"])] = {"signature": source_signature(item), "status": "ok", "records": records}

    def set_source_error(self, item: dict, error: dict) -> None:
        self.sources[str(item["id"])] = {"signature": source_signature(item), "status": "error", "error": error}

    def retain_sources(self, current_items: list[dict]) -> None:
        ids = {str(item.get("id", "")) for item in current_items}
        self.sources = {k: v for k, v in self.sources.items() if k in ids}

    def ledger_matches(self, meta: dict | None) -> bool:
        if not self.valid or not self.ledger or not meta:
            return False
        return (
            str(self.ledger.get("file_id", "")) == str(meta.get("id", ""))
            and str(self.ledger.get("md5Checksum", "")) == str(meta.get("md5Checksum", ""))
            and bool(meta.get("md5Checksum"))
        )

    def get_ledger(self) -> dict | None:
        if not self.ledger or not isinstance(self.ledger.get("data"), dict):
            return None
        return deserialize_ledger(self.ledger["data"])

    def set_ledger(self, meta: dict, ledger: dict) -> None:
        self.ledger = {
            "file_id": str(meta.get("id", "")),
            "modifiedTime": str(meta.get("modifiedTime", "")),
            "md5Checksum": str(meta.get("md5Checksum", "")),
            "size": str(meta.get("size", "")),
            "data": serialize_ledger(ledger),
        }

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_suffix(self.path.suffix + ".tmp")
        temp.write_text(json.dumps({
            "state_schema_version": STATE_SCHEMA_VERSION,
            "sources": self.sources,
            "ledger": self.ledger,
        }, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        temp.replace(self.path)
        self.valid = True
