from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from main import _read_sources, _split_historical_rows
from modules.process_parts import build_source_index
from modules.runtime_state import RuntimeState


def make_source_file(path: Path, *, order="O-1", drawing="D-1"):
    wb = Workbook()
    ws = wb.active
    headers = ["订单号", "图号", "厚度", "件数", "长(mm)", "宽(mm)", "坡口", "总净重(t)"]
    for col, value in enumerate(headers, 1):
        ws.cell(1, col, value)
    ws.append([order, drawing, 20, 3, 100, 50, "W", 0.3])
    wb.save(path)


def meta(fid, *, order="O-1", modified="2026-09-16T00:00:00Z", md5="m1"):
    return {
        "id": fid,
        "name": "汇总表.xlsx",
        "modifiedTime": modified,
        "md5Checksum": md5,
        "size": "100",
        "parent_folder_id": f"folder-{order}",
        "order_folder_id": f"folder-{order}",
        "order_container_name": order,
    }


def record(order="O-1", drawing="D-1"):
    return {
        "order_raw": order,
        "order": order,
        "drawing": drawing,
        "thickness": 20.0,
        "quantity": 3,
        "length": 100.0,
        "width": 50.0,
        "bevel": "W",
        "total_weight_t": 0.3,
        "source_file": "汇总表.xlsx",
        "order_container_name": order,
        "source_row": 2,
    }


class FakeDrive:
    def __init__(self, source_path: Path):
        self.source_path = source_path
        self.downloads = []

    def download_file(self, file_id: str, target: str):
        self.downloads.append(file_id)
        Path(target).write_bytes(self.source_path.read_bytes())
        return target


class IncrementalAcceptanceTests(unittest.TestCase):
    def test_no_change_downloads_zero_source_excels(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source_path = root / "source.xlsx"
            make_source_file(source_path)
            state_path = root / "state.json"
            state = RuntimeState(state_path)
            state.set_source_records(meta("f1"), [record()])
            state.save()
            drive = FakeDrive(source_path)
            records, errors, changes, downloads, _ = _read_sources(drive, [meta("f1")], RuntimeState(state_path), root)
            self.assertEqual(downloads, 0)
            self.assertEqual(drive.downloads, [])
            self.assertEqual(len(records), 1)
            self.assertFalse(errors)
            self.assertEqual(len(changes["unchanged"]), 1)

    def test_one_added_order_downloads_only_one(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source_path = root / "source.xlsx"
            make_source_file(source_path, order="O-2", drawing="D-2")
            state_path = root / "state.json"
            state = RuntimeState(state_path)
            state.set_source_records(meta("f1", order="O-1"), [record("O-1", "D-1")])
            state.save()
            drive = FakeDrive(source_path)
            items = [meta("f1", order="O-1"), meta("f2", order="O-2")]
            records, errors, changes, downloads, _ = _read_sources(drive, items, RuntimeState(state_path), root)
            self.assertEqual(downloads, 1)
            self.assertEqual(drive.downloads, ["f2"])
            self.assertEqual(len(changes["added"]), 1)
            self.assertEqual(len(changes["unchanged"]), 1)
            self.assertFalse(errors)
            self.assertEqual({r["order"] for r in records}, {"O-1", "O-2"})

    def test_one_modified_order_downloads_only_one_and_replaces_cache(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source_path = root / "source.xlsx"
            make_source_file(source_path, order="O-1", drawing="D-NEW")
            state_path = root / "state.json"
            state = RuntimeState(state_path)
            state.set_source_records(meta("f1"), [record("O-1", "D-OLD")])
            state.save()
            changed = meta("f1", modified="2026-09-16T01:00:00Z", md5="m2")
            drive = FakeDrive(source_path)
            runtime = RuntimeState(state_path)
            records, errors, changes, downloads, _ = _read_sources(drive, [changed], runtime, root)
            self.assertEqual(downloads, 1)
            self.assertEqual(drive.downloads, ["f1"])
            self.assertEqual(len(changes["modified"]), 1)
            self.assertFalse(errors)
            self.assertEqual([r["drawing"] for r in records], ["D-NEW"])
            cached = runtime.get_source(changed)
            self.assertEqual([r["drawing"] for r in cached["records"]], ["D-NEW"])

    def test_removed_completed_order_keeps_permanent_history(self):
        row = {**record(), "processed": 3, "remaining": 0, "pending_weight_t": 0.0, "status": "已完成"}
        completed, missing, changed = _split_historical_rows({"historical_rows": [row]}, {}, set())
        self.assertEqual(completed, [row])
        self.assertEqual(missing, [])
        self.assertEqual(changed, [])

    def test_removed_incomplete_order_is_frozen_not_deleted(self):
        row = {**record(), "processed": 1, "remaining": 2, "pending_weight_t": 0.2, "status": "部分完成"}
        completed, missing, changed = _split_historical_rows({"historical_rows": [row]}, {}, set())
        self.assertEqual(completed, [])
        self.assertEqual(missing, [row])
        self.assertEqual(changed, [])

    def test_changed_business_key_with_history_is_frozen(self):
        old = {**record(), "processed": 1, "remaining": 2, "pending_weight_t": 0.2, "status": "部分完成"}
        new = record(drawing="D-NEW")
        index = build_source_index([new])
        completed, missing, changed = _split_historical_rows({"historical_rows": [old]}, index, {"O-1"})
        self.assertEqual(completed, [])
        self.assertEqual(missing, [])
        self.assertEqual(changed, [old])


if __name__ == "__main__":
    unittest.main()
