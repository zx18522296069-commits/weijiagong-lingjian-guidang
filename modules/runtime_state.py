"""未加工零件运行时性能状态。

runtime_state/state.json 只用于加速，Google Drive 正式订单与累计台账始终是事实源。
缓存缺失、损坏、schema 不兼容或累计台账 MD5 变化时，调用方必须回退到真实 Drive 初始化。
"""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path

STATE_SCHEMA_VERSION = 1
SOURCE_CACHE_VERSION = 2


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _serialize_ledger(ledger: dict) -> dict:
    state_rows = []
    for key, value in (ledger.get("state") or {}).items():
        order, drawing, thickness, bevel = key
        state_rows.append({
            "key": [order, drawing, float(thickness), bevel],
            "value": deepcopy(value),
        })
    return {
        "existing_state": state_rows,
        "historical_rows": deepcopy(list(ledger.get("historical_rows", []))),
        "flows": deepcopy(list(ledger.get("flows", []))),
        "board_records": deepcopy(list(ledger.get("board_records", []))),
        "anomalies": deepcopy(list(ledger.get("anomalies", []))),
        "posted_boards": sorted(_text(v) for v in ledger.get("posted_boards", set()) if _text(v)),
        "posted_board_keys": sorted(
            [[_text(board), _text(fingerprint)] for board, fingerprint in ledger.get("posted_board_keys", set())]
        ),
        "legacy_posted_boards": sorted(
            _text(v) for v in ledger.get("legacy_posted_boards", set()) if _text(v)
        ),
    }


def _deserialize_ledger(payload: dict) -> dict:
    state: dict[tuple, dict] = {}
    for item in payload.get("existing_state", []):
        key = item.get("key") if isinstance(item, dict) else None
        if not isinstance(key, list) or len(key) != 4:
            raise ValueError("runtime state existing_state key 无效")
        state[(str(key[0]), str(key[1]), float(key[2]), str(key[3]))] = deepcopy(item.get("value") or {})
    return {
        "state": state,
        "historical_rows": deepcopy(list(payload.get("historical_rows", []))),
        "flows": deepcopy(list(payload.get("flows", []))),
        "board_records": deepcopy(list(payload.get("board_records", []))),
        "anomalies": deepcopy(list(payload.get("anomalies", []))),
        "posted_boards": set(_text(v) for v in payload.get("posted_boards", []) if _text(v)),
        "posted_board_keys": {
            (_text(item[0]), _text(item[1]))
            for item in payload.get("posted_board_keys", [])
            if isinstance(item, list) and len(item) == 2 and _text(item[0]) and _text(item[1])
        },
        "legacy_posted_boards": set(
            _text(v) for v in payload.get("legacy_posted_boards", []) if _text(v)
        ),
    }


class RuntimeState:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.payload = {
            "state_schema_version": STATE_SCHEMA_VERSION,
            "sources": {},
            "cumulative": None,
        }
        self.valid = False
        self.load_error = ""
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            self.load_error = "cache不存在"
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError) as exc:
            self.load_error = f"cache损坏: {exc}"
            return
        if payload.get("state_schema_version") != STATE_SCHEMA_VERSION:
            self.load_error = "schema版本不兼容"
            return
        if not isinstance(payload.get("sources", {}), dict):
            self.load_error = "sources结构无效"
            return
        cumulative = payload.get("cumulative")
        if cumulative is not None and not isinstance(cumulative, dict):
            self.load_error = "cumulative结构无效"
            return
        self.payload = payload
        self.valid = True

    @staticmethod
    def _meta_signature(meta: dict) -> dict:
        return {
            "file_id": _text(meta.get("id") or meta.get("file_id")),
            "modifiedTime": _text(meta.get("modifiedTime")),
            "md5Checksum": _text(meta.get("md5Checksum")).lower(),
            "size": _text(meta.get("size")),
        }

    def get_cumulative_ledger(self, drive_meta: dict | None) -> dict | None:
        if not self.valid or not drive_meta:
            return None
        cumulative = self.payload.get("cumulative")
        if not isinstance(cumulative, dict):
            return None
        if cumulative.get("metadata") != self._meta_signature(drive_meta):
            return None
        ledger = cumulative.get("ledger")
        if not isinstance(ledger, dict):
            return None
        try:
            return _deserialize_ledger(ledger)
        except (TypeError, ValueError, KeyError):
            return None

    def set_cumulative_ledger(self, drive_meta: dict, ledger: dict) -> None:
        self.payload["cumulative"] = {
            "metadata": self._meta_signature(drive_meta),
            "ledger": _serialize_ledger(ledger),
        }
        self.valid = True

    def set_source_entries(self, entries: dict) -> None:
        self.payload["sources"] = deepcopy(entries if isinstance(entries, dict) else {})
        self.valid = True

    def restore_source_cache(self, source_cache_path: str | Path) -> bool:
        """当独立 source cache 缺失时，用 state.json 中的同一份解析结果恢复性能缓存。"""
        path = Path(source_cache_path)
        if path.exists() or not self.valid:
            return False
        sources = self.payload.get("sources")
        if not isinstance(sources, dict) or not sources:
            return False
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps({"version": SOURCE_CACHE_VERSION, "entries": sources}, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )
        return True

    def capture_source_cache(self, source_cache_path: str | Path) -> None:
        path = Path(source_cache_path)
        if not path.exists():
            return
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            return
        entries = payload.get("entries")
        if isinstance(entries, dict):
            self.set_source_entries(entries)

    def save(self) -> None:
        self.payload["state_schema_version"] = STATE_SCHEMA_VERSION
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_suffix(self.path.suffix + ".tmp")
        temp.write_text(
            json.dumps(self.payload, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )
        temp.replace(self.path)
        self.valid = True
