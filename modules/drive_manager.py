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
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class DriveManager:
    def __init__(self):
        self.root_folder_id = os.getenv("DRIVE_ROOT_FOLDER_ID", "")
        self.working_folder_id = os.getenv("WORKING_FOLDER_ID", "")
        self.split_folder_id = os.getenv("SPLIT_FOLDER_ID", "")
        self.archive_folder_id = os.getenv("ARCHIVE_FOLDER_ID", "")
        cache_root = os.getenv("DRIVE_DOWNLOAD_CACHE_DIR", "").strip()
        self.download_cache_root = Path(cache_root) if cache_root else None
        self.service = None

    def check_config(self) -> bool:
        return bool(
            check_auth_config()
            and self.root_folder_id
            and self.working_folder_id
            and self.split_folder_id
            and self.archive_folder_id
        )

    def connect(self):
        credentials = build_credentials()
        self.service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        return self.service

    def _service(self):
        return self.service or self.connect()

    def list_children(self, folder_id: str) -> list[dict]:
        """只列指定目录的直接子项，不递归；只读取 metadata，不下载正文。"""
        service = self._service()
        query = f"'{folder_id}' in parents and trashed=false"
        fields = (
            "nextPageToken,files(id,name,mimeType,parents,modifiedTime,md5Checksum,size,"
            "shortcutDetails(targetId,targetMimeType))"
        )
        items: list[dict] = []
        page_token = None
        while True:
            result = service.files().list(
                q=query,
                fields=fields,
                pageToken=page_token,
                pageSize=1000,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()
            items.extend(result.get("files", []))
            page_token = result.get("nextPageToken")
            if not page_token:
                break
        return items

    def get_file(self, file_id: str, fields: str = "id,name,mimeType,parents,modifiedTime,md5Checksum,size") -> dict:
        return self._service().files().get(
            fileId=file_id,
            fields=fields,
            supportsAllDrives=True,
        ).execute()

    def resolve_folder_item(self, item: dict) -> tuple[str, str] | None:
        """把普通文件夹或指向文件夹的快捷方式解析成 (folder_id, display_name)。"""
        if item.get("mimeType") == FOLDER_MIME:
            return item["id"], item.get("name", "")
        if item.get("mimeType") == SHORTCUT_MIME:
            details = item.get("shortcutDetails") or {}
            if details.get("targetMimeType") == FOLDER_MIME and details.get("targetId"):
                return details["targetId"], item.get("name", "")
        return None

    def list_order_containers(self) -> list[dict]:
        """返回“正在加工”根目录中的全部订单文件夹/文件夹快捷方式。"""
        result: list[dict] = []
        for item in self.list_children(self.working_folder_id):
            resolved = self.resolve_folder_item(item)
            if not resolved:
                continue
            folder_id, container_name = resolved
            result.append({
                "id": folder_id,
                "name": container_name,
                "source_item_id": item.get("id", folder_id),
            })
        return result

    @staticmethod
    def _is_excel_file(item: dict) -> bool:
        name = str(item.get("name", ""))
        return item.get("mimeType") not in {FOLDER_MIME, SHORTCUT_MIME} and name.lower().endswith((".xlsx", ".xlsm", ".xls"))

    def _order_excel_candidates(self, folder_id: str) -> list[dict]:
        """
        收集订单目录内可作为正式汇总表候选的 Excel。

        安全规则：仅扫描订单目录本层和下一层子文件夹，不继续深层递归；
        这样兼容“订单/全/模版.xlsm”这类实际目录，同时避免误抓更深层附件。
        """
        direct_children = self.list_children(folder_id)
        excel_files = []
        for item in direct_children:
            if not self._is_excel_file(item):
                continue
            candidate = dict(item)
            candidate["parent_folder_id"] = folder_id
            excel_files.append(candidate)

        for item in direct_children:
            resolved = self.resolve_folder_item(item)
            if not resolved:
                continue
            child_folder_id, child_folder_name = resolved
            for nested in self.list_children(child_folder_id):
                if not self._is_excel_file(nested):
                    continue
                candidate = dict(nested)
                candidate["source_subfolder_name"] = child_folder_name
                candidate["source_subfolder_id"] = child_folder_id
                candidate["parent_folder_id"] = child_folder_id
                excel_files.append(candidate)
        return excel_files

    def list_order_source_files(self) -> list[dict]:
        """
        扫描“正在加工”订单目录；支持文件夹快捷方式和安全的一层子目录。

        正式订单源允许两种已验证命名：
        1. 文件名包含“汇总表”；
        2. 若不存在“汇总表”命名文件，则使用文件名包含“模板”或“模版”的 Excel，
           其工作簿内部必须仍由读取层验证存在正式“汇总表”结构。

        同一订单同一优先级若出现多个候选文件则阻断，禁止猜测使用哪一个。
        不把正式台账/当前待加工表当成订单源。
        """
        result: list[dict] = []
        for container in self.list_order_containers():
            folder_id = container["id"]
            container_name = container["name"]
            excel_children = self._order_excel_candidates(folder_id)

            preferred = [child for child in excel_children if "汇总表" in str(child.get("name", ""))]
            fallback = [
                child for child in excel_children
                if "模板" in str(child.get("name", "")) or "模版" in str(child.get("name", ""))
            ]
            candidates = preferred if preferred else fallback

            if len(candidates) > 1:
                names = [
                    f"{child.get('source_subfolder_name') + '/' if child.get('source_subfolder_name') else ''}{child.get('name', '')}"
                    for child in candidates
                ]
                raise RuntimeError(
                    f"订单目录存在多个原始汇总表候选，禁止自动猜测: {container_name} -> {names}"
                )
            if not candidates:
                continue

            child = candidates[0]
            result.append({
                **child,
                "parent_folder_id": child.get("parent_folder_id") or folder_id,
                "order_container_name": container_name,
                "order_folder_id": folder_id,
            })
        return result

    def list_pending_split_files(self) -> list[dict]:
        """
        只扫描“拆图结果”根目录直接子文件；不递归进入“已录入数量”。
        仅返回 *_完成.xlsx / *_完成.xlsm 等待入账文件。
        """
        result: list[dict] = []
        for item in self.list_children(self.split_folder_id):
            name = str(item.get("name", ""))
            lower = name.lower()
            if item.get("mimeType") in {FOLDER_MIME, SHORTCUT_MIME}:
                continue
            if "_完成" not in name:
                continue
            if not lower.endswith((".xlsx", ".xlsm", ".xls")):
                continue
            result.append(item)
        return result

    @staticmethod
    def board_id_from_filename(filename: str) -> str:
        """
        板材号取完成文件的原始板号，并保留小序号。

        例如：#2323_完成.xlsx -> #2323；
        #2323_完成 (1).xlsx -> #2323-1。
        文件系统追加的 (1)、（1）统一视为板材小号，不能与原板重复入账。
        """
        name = Path(filename).stem
        if "_完成" not in name:
            raise ValueError(f"不是完成文件: {filename}")
        board_id, suffix = name.split("_完成", 1)
        if not board_id:
            raise ValueError(f"无法从文件名提取板材号: {filename}")
        match = re.fullmatch(r"\s*[（(]\s*(\d+)\s*[)）]\s*", suffix)
        if match:
            return f"{board_id}-{match.group(1)}"
        return board_id

    def _cache_dir_for_file(self, file_id: str) -> Path | None:
        if self.download_cache_root is None:
            return None
        safe_id = re.sub(r"[^A-Za-z0-9_-]", "_", file_id)
        return self.download_cache_root / safe_id

    def _cache_path_for_metadata(self, file_id: str, metadata: dict) -> Path | None:
        cache_dir = self._cache_dir_for_file(file_id)
        if cache_dir is None:
            return None
        signature = "|".join(
            [
                file_id,
                str(metadata.get("modifiedTime", "")),
                str(metadata.get("md5Checksum", "")),
                str(metadata.get("size", "")),
            ]
        )
        digest = hashlib.sha256(signature.encode("utf-8")).hexdigest()
        return cache_dir / f"{digest}.bin"

    def _invalidate_download_cache(self, file_id: str) -> None:
        cache_dir = self._cache_dir_for_file(file_id)
        if cache_dir and cache_dir.exists():
            shutil.rmtree(cache_dir, ignore_errors=True)

    def download_file(self, file_id: str, save_path: str) -> str:
        """下载 Drive 文件；配置缓存时优先复用 file id + modifiedTime + MD5 + size 一致的本地正文。"""
        service = self._service()
        target = Path(save_path)
        target.parent.mkdir(parents=True, exist_ok=True)

        cache_path = None
        if self.download_cache_root is not None:
            metadata = self.get_file(file_id, fields="id,modifiedTime,md5Checksum,size")
            cache_path = self._cache_path_for_metadata(file_id, metadata)
            if cache_path and cache_path.exists():
                expected_size = metadata.get("size")
                if expected_size in (None, "") or cache_path.stat().st_size == int(expected_size):
                    shutil.copy2(cache_path, target)
                    return str(target)
                cache_path.unlink(missing_ok=True)

        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        with io.FileIO(target, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()

        if cache_path is not None:
            cache_dir = cache_path.parent
            cache_dir.mkdir(parents=True, exist_ok=True)
            for old in cache_dir.glob("*.bin"):
                if old != cache_path:
                    old.unlink(missing_ok=True)
            shutil.copy2(target, cache_path)
        return str(target)

    def upload_file(self, file_path: str, target_folder_id: str) -> dict:
        metadata = {"name": os.path.basename(file_path), "parents": [target_folder_id]}
        media = MediaFileUpload(file_path, resumable=False)
        return self._service().files().create(
            body=metadata,
            media_body=media,
            fields="id,name,parents,modifiedTime,md5Checksum,size",
            supportsAllDrives=True,
        ).execute()

    def update_file_content(self, file_id: str, file_path: str) -> dict:
        """覆盖既有 Drive 文件内容，保留原文件 ID/位置。"""
        self._invalidate_download_cache(file_id)
        media = MediaFileUpload(file_path, resumable=False)
        return self._service().files().update(
            fileId=file_id,
            media_body=media,
            fields="id,name,parents,modifiedTime,md5Checksum,size",
            supportsAllDrives=True,
        ).execute()

    def move_file(self, file_id: str, target_folder_id: str) -> dict:
        service = self._service()
        file_info = service.files().get(
            fileId=file_id,
            fields="parents",
            supportsAllDrives=True,
        ).execute()
        previous = ",".join(file_info.get("parents", []))
        return service.files().update(
            fileId=file_id,
            addParents=target_folder_id,
            removeParents=previous,
            fields="id,name,parents,modifiedTime",
            supportsAllDrives=True,
        ).execute()
