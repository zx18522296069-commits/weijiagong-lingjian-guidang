from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from modules.runtime_state import RuntimeState, STATE_SCHEMA_VERSION


class RuntimeStateTests(unittest.TestCase):
    def _meta(self, md5="abc"):
        return {
            "id": "ledger-1",
            "modifiedTime": "2026-09-16T10:00:00Z",
            "md5Checksum": md5,
            "size": "12345",
        }

    def _ledger(self):
        key = ("ORDER-1", "A1", 20.0, "W")
        return {
            "state": {key: {"processed": 2, "board_sources": "#1×2"}},
            "historical_rows": [{
                "order": "ORDER-1", "drawing": "A1", "thickness": 20.0, "bevel": "W",
                "quantity": 5, "processed": 2, "remaining": 3,
            }],
            "flows": [{"板材号": "#1", "内容指纹": "fp1", "本张板加工数量": 2}],
            "board_records": [{"板材号": "#1", "内容指纹": "fp1", "状态": "已核验并入账"}],
            "anomalies": [{"板材号": "#9", "异常类型": "阻断入账"}],
            "posted_boards": {"#1"},
            "posted_board_keys": {("#1", "fp1")},
            "legacy_posted_boards": set(),
        }

    def test_cumulative_roundtrip_and_md5_guard(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "runtime_state" / "state.json"
            state = RuntimeState(path)
            state.set_cumulative_ledger(self._meta(), self._ledger())
            state.save()

            restored = RuntimeState(path)
            self.assertTrue(restored.valid)
            ledger = restored.get_cumulative_ledger(self._meta())
            self.assertIsNotNone(ledger)
            self.assertEqual(ledger["state"][("ORDER-1", "A1", 20.0, "W")]["processed"], 2)
            self.assertIn(("#1", "fp1"), ledger["posted_board_keys"])
            self.assertIsNone(restored.get_cumulative_ledger(self._meta("changed")))

    def test_corrupt_or_schema_mismatch_falls_back(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "state.json"
            path.write_text("{bad json", encoding="utf-8")
            corrupt = RuntimeState(path)
            self.assertFalse(corrupt.valid)
            self.assertIsNone(corrupt.get_cumulative_ledger(self._meta()))

            path.write_text(json.dumps({"state_schema_version": STATE_SCHEMA_VERSION + 1}), encoding="utf-8")
            incompatible = RuntimeState(path)
            self.assertFalse(incompatible.valid)
            self.assertIsNone(incompatible.get_cumulative_ledger(self._meta()))

    def test_sources_are_persisted_and_can_restore_source_cache(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            state_path = root / "runtime_state" / "state.json"
            source_path = root / ".cache" / "source_records.json"
            entries = {
                "source-1": {
                    "signature": {
                        "file_id": "source-1",
                        "modifiedTime": "2026-09-16T01:00:00Z",
                        "md5Checksum": "md5",
                        "parent_folder_id": "folder-1",
                        "name": "汇总表.xlsx",
                        "order_container_name": "ORDER-1",
                    },
                    "status": "ok",
                    "records": [{"order": "ORDER-1", "drawing": "A1"}],
                }
            }
            state = RuntimeState(state_path)
            state.set_source_entries(entries)
            state.save()

            restored = RuntimeState(state_path)
            self.assertTrue(restored.restore_source_cache(source_path))
            payload = json.loads(source_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["version"], 2)
            self.assertEqual(payload["entries"], entries)


if __name__ == "__main__":
    unittest.main()
