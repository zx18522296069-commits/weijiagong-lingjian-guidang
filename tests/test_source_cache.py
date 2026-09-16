from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from modules.source_cache import CACHE_VERSION, SourceRecordCache


class SourceRecordCacheTests(unittest.TestCase):
    def _item(
        self,
        modified="2026-09-16T01:00:00.000Z",
        *,
        md5="abc123",
        parent="order-folder",
    ):
        return {
            "id": "source-1",
            "name": "订单汇总表.xlsx",
            "modifiedTime": modified,
            "md5Checksum": md5,
            "size": "1234",
            "parent_folder_id": parent,
            "order_container_name": "190.26-09-11 THP10-8000J-0911",
        }

    def test_unchanged_source_reuses_records(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            cache = SourceRecordCache(path)
            item = self._item()
            records = [{"order": "THP10-8000J-0911", "drawing": "2310", "thickness": 80.0}]
            cache.store_records(item, records)
            cache.save()

            restored = SourceRecordCache(path)
            hit = restored.get(item)
            self.assertIsNotNone(hit)
            self.assertEqual(hit["records"], records)

    def test_modified_source_invalidates_records(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            cache = SourceRecordCache(path)
            cache.store_records(self._item(), [{"order": "OLD"}])
            cache.save()

            restored = SourceRecordCache(path)
            self.assertIsNone(restored.get(self._item("2026-09-16T02:00:00.000Z")))

    def test_md5_change_invalidates_records(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            cache = SourceRecordCache(path)
            cache.store_records(self._item(), [{"order": "OLD"}])
            cache.save()
            self.assertIsNone(SourceRecordCache(path).get(self._item(md5="different")))

    def test_parent_move_invalidates_records(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            cache = SourceRecordCache(path)
            cache.store_records(self._item(), [{"order": "OLD"}])
            cache.save()
            self.assertIsNone(SourceRecordCache(path).get(self._item(parent="nested-folder")))

    def test_removed_source_is_pruned(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            cache = SourceRecordCache(path)
            cache.store_records(self._item(), [{"order": "OLD"}])
            cache.retain([])
            cache.save()
            self.assertEqual(SourceRecordCache(path).entries, {})

    def test_v1_cache_is_reused_once_when_legacy_signature_is_unchanged(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            item = self._item()
            legacy = {
                "version": 1,
                "entries": {
                    "source-1": {
                        "signature": {
                            "modifiedTime": item["modifiedTime"],
                            "size": item["size"],
                            "name": item["name"],
                            "order_container_name": item["order_container_name"],
                        },
                        "status": "ok",
                        "records": [{"order": "CACHED"}],
                    }
                },
            }
            path.write_text(json.dumps(legacy), encoding="utf-8")

            cache = SourceRecordCache(path)
            hit = cache.get(item)
            self.assertIsNotNone(hit)
            self.assertEqual(hit["records"][0]["order"], "CACHED")
            self.assertEqual(hit["signature"]["md5Checksum"], "abc123")
            self.assertEqual(hit["signature"]["parent_folder_id"], "order-folder")
            cache.save()
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["version"], CACHE_VERSION)

    def test_v1_cache_is_not_reused_if_legacy_metadata_changed(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            item = self._item()
            legacy = {
                "version": 1,
                "entries": {
                    "source-1": {
                        "signature": {
                            "modifiedTime": "older",
                            "size": item["size"],
                            "name": item["name"],
                            "order_container_name": item["order_container_name"],
                        },
                        "status": "ok",
                        "records": [{"order": "STALE"}],
                    }
                },
            }
            path.write_text(json.dumps(legacy), encoding="utf-8")
            self.assertIsNone(SourceRecordCache(path).get(item))


if __name__ == "__main__":
    unittest.main()
