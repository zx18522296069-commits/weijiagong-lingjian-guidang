from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from modules.source_cache import SourceRecordCache


class SourceRecordCacheTests(unittest.TestCase):
    def _item(self, modified="2026-09-16T01:00:00.000Z"):
        return {
            "id": "source-1",
            "name": "订单汇总表.xlsx",
            "modifiedTime": modified,
            "size": "1234",
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

    def test_removed_source_is_pruned(self):
        with tempfile.TemporaryDirectory() as temp_name:
            path = Path(temp_name) / "source_records.json"
            cache = SourceRecordCache(path)
            cache.store_records(self._item(), [{"order": "OLD"}])
            cache.retain([])
            cache.save()
            self.assertEqual(SourceRecordCache(path).entries, {})


if __name__ == "__main__":
    unittest.main()
