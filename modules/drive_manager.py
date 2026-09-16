"""Google Drive 连接与正式目录扫描模块。"""
from __future__ import annotations

import hashlib
import io
import os
import re
import shutil
from pathlib import Path

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from modules.google_auth import build_credentials, check_auth_config

FOLDER_MIME = "application/vnd.google-apps.folder"
SHORTCUT_MIME = "application/vnd.google-apps.shortcut"


class DriveManager:
    def __init__(self):
        self.root_folder_id = os.getenv("DRIVE_ROOT_FOLDER_ID", "")
        self.working_folder_id = os.getenv("WORKING_FOLDER_ID", "")
        self.split_folder_id = os.getenv("SPLIT_FOLDER_ID", "")
        self.archive_folder_id = os.getenv("ARCHIVE_FOLDER_ID", "")
        cache_root = os.getenv("DRIVE_DOWNLOAD_CACHE_DIR", "").strip()
        self.download_cache_root = Path(cache_root) if cache_root else None
        self.service = None
        self.stats = {"downloads": 0, "uploads": 0, "metadata_reads": 0}

    def check_config(self) -> bool:
        return bool(check_auth_config() and self.root_folder_id and self.working_folder_id and self.split_folder_id and self.archive_folder_id)

    def connect(self):
        self.service = build("drive", "v3", credentials=build_credentials(), cache_discovery=False)
        return self.service

    def _service(self): return self.service or self.connect()

    def list_children(self, folder_id: str) -> list[dict]:
        query = f"'{folder_id}' in parents and trashed=false"
        fields = "nextPageToken,files(id,name,mimeType,parents,modifiedTime,size,md5Checksum,shortcutDetails(targetId,targetMimeType))"
        items, page_token = [], None
        while True:
            result = self._service().files().list(q=query, fields=fields, pageToken=page_token, pageSize=1000,
                supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
            items.extend(result.get("files", [])); page_token = result.get("nextPageToken")
            if not page_token: break
        self.stats["metadata_reads"] += 1
        return items

    def get_file(self, file_id: str, fields: str = "id,name,mimeType,parents,modifiedTime,size,md5Checksum") -> dict:
        self.stats["metadata_reads"] += 1
        return self._service().files().get(fileId=file_id, fields=fields, supportsAllDrives=True).execute()

    def resolve_folder_item(self, item: dict) -> tuple[str, str] | None:
        if item.get("mimeType") == FOLDER_MIME: return item["id"], item.get("name", "")
        if item.get("mimeType") == SHORTCUT_MIME:
            details = item.get("shortcutDetails") or {}
            if details.get("targetMimeType") == FOLDER_MIME and details.get("targetId"):
                return details["targetId"], item.get("name", "")
        return None

    def list_order_source_files(self) -> list[dict]:
        result = []
        for item in self.list_children(self.working_folder_id):
            resolved = self.resolve_folder_item(item)
            if not resolved: continue
            folder_id, container_name = resolved; excel_children = []
            for child in self.list_children(folder_id):
                name = str(child.get("name", "")); lower = name.lower()
                if child.get("mimeType") == FOLDER_MIME or not lower.endswith((".xlsx", ".xlsm", ".xls")): continue
                excel_children.append(child)
            preferred = [x for x in excel_children if "汇总表" in str(x.get("name", ""))]
            fallback = [x for x in excel_children if "模板" in str(x.get("name", ""))]
            candidates = preferred if preferred else fallback
            if len(candidates) > 1: raise RuntimeError(f"订单目录存在多个原始汇总表候选，禁止自动猜测: {container_name} -> {[x.get('name') for x in candidates]}")
            if not candidates: continue
            child = candidates[0]
            result.append({**child, "order_container_name": container_name, "order_folder_id": folder_id, "parent_folder_id": folder_id})
        return result

    def list_pending_split_files(self) -> list[dict]:
        result = []
        for item in self.list_children(self.split_folder_id):
            name = str(item.get("name", "")); lower = name.lower()
            if item.get("mimeType") == FOLDER_MIME or "_完成" not in name: continue
            if lower.endswith((".xlsx", ".xlsm", ".xls")): result.append(item)
        return result

    @staticmethod
    def board_id_from_filename(filename: str) -> str:
        name = Path(filename).stem
        if "_完成" not in name: raise ValueError(f"不是完成文件: {filename}")
        board_id, suffix = name.split("_完成", 1)
        if not board_id: raise ValueError(f"无法从文件名提取板材号: {filename}")
        match = re.fullmatch(r"\s*[（(]\s*(\d+)\s*[)）]\s*", suffix)
        return f"{board_id}-{match.group(1)}" if match else board_id

    def _cache_path(self, file_id: str, meta: dict) -> Path | None:
        if self.download_cache_root is None: return None
        safe = re.sub(r"[^A-Za-z0-9_-]", "_", file_id)
        signature = "|".join([file_id, str(meta.get("modifiedTime", "")), str(meta.get("md5Checksum", "")), str(meta.get("size", ""))])
        return self.download_cache_root / safe / (hashlib.sha256(signature.encode()).hexdigest() + ".bin")

    def _invalidate_download_cache(self, file_id: str) -> None:
        if self.download_cache_root is None: return
        safe = re.sub(r"[^A-Za-z0-9_-]", "_", file_id); shutil.rmtree(self.download_cache_root / safe, ignore_errors=True)

    def download_file(self, file_id: str, save_path: str) -> str:
        target = Path(save_path); target.parent.mkdir(parents=True, exist_ok=True)
        cache_path = None
        if self.download_cache_root is not None:
            meta = self.get_file(file_id, fields="id,modifiedTime,size,md5Checksum")
            cache_path = self._cache_path(file_id, meta)
            if cache_path and cache_path.exists() and (not meta.get("size") or cache_path.stat().st_size == int(meta["size"])):
                shutil.copy2(cache_path, target); return str(target)
        request = self._service().files().get_media(fileId=file_id, supportsAllDrives=True)
        with io.FileIO(target, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request); done = False
            while not done: _, done = downloader.next_chunk()
        self.stats["downloads"] += 1
        if cache_path:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            for old in cache_path.parent.glob("*.bin"):
                if old != cache_path: old.unlink(missing_ok=True)
            shutil.copy2(target, cache_path)
        return str(target)

    @staticmethod
    def local_hashes(path: str | Path) -> dict:
        sha256 = hashlib.sha256(); md5 = hashlib.md5(usedforsecurity=False); size = 0
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                sha256.update(chunk); md5.update(chunk); size += len(chunk)
        return {"sha256": sha256.hexdigest(), "md5": md5.hexdigest(), "size": size}

    def upload_file(self, file_path: str, target_folder_id: str) -> dict:
        result = self._service().files().create(body={"name": os.path.basename(file_path), "parents": [target_folder_id]},
            media_body=MediaFileUpload(file_path, resumable=False), fields="id,name,parents,modifiedTime,size,md5Checksum", supportsAllDrives=True).execute()
        self.stats["uploads"] += 1; return result

    def update_file_content(self, file_id: str, file_path: str) -> dict:
        self._invalidate_download_cache(file_id)
        result = self._service().files().update(fileId=file_id, media_body=MediaFileUpload(file_path, resumable=False),
            fields="id,name,parents,modifiedTime,size,md5Checksum", supportsAllDrives=True).execute()
        self.stats["uploads"] += 1; return result

    def verify_remote_hash(self, file_id: str, local_path: str | Path) -> dict:
        local = self.local_hashes(local_path); remote = self.get_file(file_id, fields="id,name,parents,modifiedTime,size,md5Checksum")
        if not remote.get("md5Checksum"): raise RuntimeError(f"Drive未返回MD5，无法确认上传完整性: {remote.get('name') or file_id}")
        if str(remote["md5Checksum"]).lower() != local["md5"].lower(): raise RuntimeError(f"Drive MD5校验失败: {remote.get('name') or file_id}")
        if remote.get("size") not in (None, "") and int(remote["size"]) != local["size"]: raise RuntimeError(f"Drive size校验失败: {remote.get('name') or file_id}")
        return {**remote, "local_sha256": local["sha256"], "local_md5": local["md5"]}

    def move_file(self, file_id: str, target_folder_id: str) -> dict:
        info = self._service().files().get(fileId=file_id, fields="parents", supportsAllDrives=True).execute(); previous = ",".join(info.get("parents", []))
        return self._service().files().update(fileId=file_id, addParents=target_folder_id, removeParents=previous,
            fields="id,name,parents,modifiedTime", supportsAllDrives=True).execute()
