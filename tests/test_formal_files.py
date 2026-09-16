import tempfile
import unittest
from pathlib import Path

from modules.formal_files import (
    CUMULATIVE_NAME,
    PENDING_NAME,
    discover_formal_files,
    empty_ledger,
    find_unique_working_file,
    write_formal_file,
)


class FakeDrive:
    working_folder_id = "working"

    def __init__(self, children=None):
        self.children = list(children or [])
        self.updated = []
        self.uploaded = []

    def list_children(self, folder_id):
        self.assert_folder = folder_id
        return list(self.children)

    def update_file_content(self, file_id, file_path):
        self.updated.append((file_id, file_path))
        return {"id": file_id, "name": Path(file_path).name}

    def upload_file(self, file_path, folder_id):
        self.uploaded.append((file_path, folder_id))
        return {"id": "new-id", "name": Path(file_path).name}


class FormalFilesTest(unittest.TestCase):
    def test_discovers_both_formal_files_by_exact_name(self):
        drive = FakeDrive([
            {"id": "c1", "name": CUMULATIVE_NAME, "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            {"id": "p1", "name": PENDING_NAME, "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            {"id": "folder", "name": "订单A", "mimeType": "application/vnd.google-apps.folder"},
        ])
        found = discover_formal_files(drive)
        self.assertEqual(found["cumulative"]["id"], "c1")
        self.assertEqual(found["pending"]["id"], "p1")

    def test_missing_file_returns_none(self):
        drive = FakeDrive([])
        self.assertIsNone(find_unique_working_file(drive, CUMULATIVE_NAME))

    def test_duplicate_formal_file_is_blocked(self):
        drive = FakeDrive([
            {"id": "a", "name": CUMULATIVE_NAME, "mimeType": "x"},
            {"id": "b", "name": CUMULATIVE_NAME, "mimeType": "x"},
        ])
        with self.assertRaisesRegex(RuntimeError, "多个同名正式文件"):
            find_unique_working_file(drive, CUMULATIVE_NAME)

    def test_empty_ledger_has_all_permanent_sections(self):
        ledger = empty_ledger()
        self.assertEqual(ledger["state"], {})
        self.assertEqual(ledger["historical_rows"], [])
        self.assertEqual(ledger["flows"], [])
        self.assertEqual(ledger["board_records"], [])
        self.assertEqual(ledger["anomalies"], [])
        self.assertEqual(ledger["posted_boards"], set())
        self.assertEqual(ledger["posted_board_keys"], set())

    def test_write_updates_existing_or_uploads_missing(self):
        drive = FakeDrive([])
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / CUMULATIVE_NAME
            path.write_bytes(b"x")
            existing = write_formal_file(drive, {"id": "old-id"}, path)
            self.assertEqual(existing["id"], "old-id")
            self.assertEqual(len(drive.updated), 1)

            created = write_formal_file(drive, None, path)
            self.assertEqual(created["id"], "new-id")
            self.assertEqual(drive.uploaded[0][1], "working")


if __name__ == "__main__":
    unittest.main()
