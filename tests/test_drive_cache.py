from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from modules.drive_manager import DriveManager


class _Execute:
    def __init__(self, value):
        self.value = value

    def execute(self):
        return self.value


class _FakeFiles:
    def __init__(self):
        self.metadata = {"id": "file-1", "modifiedTime": "2026-09-16T01:00:00.000Z", "size": "3"}
        self.payload = b"abc"
        self.media_calls = 0

    def get(self, **kwargs):
        return _Execute(dict(self.metadata))

    def get_media(self, **kwargs):
        self.media_calls += 1
        return self.payload


class _FakeService:
    def __init__(self):
        self.files_api = _FakeFiles()

    def files(self):
        return self.files_api


class _FakeDownloader:
    def __init__(self, handle, request):
        self.handle = handle
        self.request = request
        self.done = False

    def next_chunk(self):
        if not self.done:
            self.handle.write(self.request)
            self.done = True
        return None, True


class DriveCacheTests(unittest.TestCase):
    def test_same_file_id_and_modified_time_reuses_cache(self):
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            manager = DriveManager()
            manager.download_cache_root = root / "cache"
            manager.service = _FakeService()

            first = root / "first.xlsx"
            second = root / "second.xlsx"

            with patch("modules.drive_manager.MediaIoBaseDownload", _FakeDownloader):
                manager.download_file("file-1", str(first))
                self.assertEqual(first.read_bytes(), b"abc")
                self.assertEqual(manager.service.files_api.media_calls, 1)

                # 远端内容变量即使变化，只要 Drive 的 modifiedTime/size 未变化，
                # 第二次也必须复用已验证缓存，不再次拉文件体。
                manager.service.files_api.payload = b"zzz"
                manager.download_file("file-1", str(second))
                self.assertEqual(second.read_bytes(), b"abc")
                self.assertEqual(manager.service.files_api.media_calls, 1)

    def test_modified_time_change_forces_redownload(self):
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            manager = DriveManager()
            manager.download_cache_root = root / "cache"
            manager.service = _FakeService()

            first = root / "first.xlsx"
            changed = root / "changed.xlsx"

            with patch("modules.drive_manager.MediaIoBaseDownload", _FakeDownloader):
                manager.download_file("file-1", str(first))
                manager.service.files_api.metadata["modifiedTime"] = "2026-09-16T02:00:00.000Z"
                manager.service.files_api.payload = b"xyz"
                manager.download_file("file-1", str(changed))

            self.assertEqual(changed.read_bytes(), b"xyz")
            self.assertEqual(manager.service.files_api.media_calls, 2)


if __name__ == "__main__":
    unittest.main()
