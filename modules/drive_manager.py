"""Google Drive 连接与正式目录扫描模块。"""

from __future__ import annotations

import io
import json
import os
from pathlib import Path
from typing import Iterable

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_MIME = "application/vnd.google-apps.folder"
SHORTCUT_MIME = "application/vnd.google-apps.shortcut"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class DriveManager:
    def __init__(self):
        self.root_folder_id = os.getenv("DRIVE_ROOT_FOLDER_ID", "")
        self.working_folder_id = os.getenv("WORKING_FOLDER_ID", "")
        self.split_folder_id = os.getenv("SPLIT_FOLDER_ID", "")
        self.archive_folder_id = os.getenv("ARCHIVE_FOLDER_ID", "")
        self.credentials_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
        self.service = None

    def check_config(self) -> bool:
        return bool(
            self.credentials_json
            and self.root_folder_id
            and self.working_folder_id
            and self.split_folder_id
            and self.archive_folder_id
        )

    def connect(self):
        if not self.credentials_json:
            raise RuntimeError("缺少 GOOGLE_SERVICE_ACCOUNT_JSON")
        info = json.loads(self.credentials_json)
        credentials = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        self.service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        return self.service

    def _service(self):
        return self.service or self.connect()

    def list_children(self, folder_id: str) -> list[dict]:
        """只列指定目录的直接子项，不递归。"""
        service = self._service()
        query = f"'{folder_id}' in parents and trashed=false"
        fields = (
            "nextPageToken,files(id,name,mimeType,parents,modifiedTime,size,"
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

    def get_file(self, file_id: str, fields: str = "id,name,mimeType,parents,modifiedTime,size") -> dict:
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

    def list_order_source_files(self) -> list[dict]:
        """
        扫描“正在加工”直接子项；支持订单文件夹和文件夹快捷方式。
        每个订单只返回其目录下名称含“汇总表”的 Excel 文件。
        不把正式台账/当前待加工表当成订单源。
        """
        result: list[dict] = []
        for item in self.list_children(self.working_folder_id):
            resolved = self.resolve_folder_item(item)
            if not resolved:
                continue
            folder_id, container_name = resolved
            for child in self.list_children(folder_id):
                name = str(child.get("name", ""))
                lower = name.lower()
                if "汇总表" not in name or not lower.endswith((".xlsx", ".xlsm", ".xls")):
                    continue
                result.append({
                    **child,
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
            if item.get("mimeType") == FOLDER_MIME:
                continue
            if "_完成" not in name:
                continue
            if not lower.endswith((".xlsx", ".xlsm", ".xls")):
                continue
            result.append(item)
        return result

    @staticmethod
    def board_id_from_filename(filename: str) -> str:
        """文件名中 `_完成` 前的完整字符串就是板材号，保留 #、废 等字符。"""
        name = Path(filename).stem
        if "_完成" not in name:
            raise ValueError(f"不是完成文件: {filename}")
        board_id = name.split("_完成", 1)[0]
        if not board_id:
            raise ValueError(f"无法从文件名提取板材号: {filename}")
        return board_id

    def download_file(self, file_id: str, save_path: str) -> str:
        service = self._service()
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        with io.FileIO(save_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return save_path

    def upload_file(self, file_path: str, target_folder_id: str) -> dict:
        metadata = {"name": os.path.basename(file_path), "parents": [target_folder_id]}
        media = MediaFileUpload(file_path, resumable=False)
        return self._service().files().create(
            body=metadata,
            media_body=media,
            fields="id,name,parents,modifiedTime,size",
            supportsAllDrives=True,
        ).execute()

    def update_file_content(self, file_id: str, file_path: str) -> dict:
        """覆盖既有 Drive 文件内容，保留原文件 ID/位置。"""
        media = MediaFileUpload(file_path, resumable=False)
        return self._service().files().update(
            fileId=file_id,
            media_body=media,
            fields="id,name,parents,modifiedTime,size",
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
