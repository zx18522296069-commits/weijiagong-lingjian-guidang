"""
Google Drive连接模块

用于GitHub Actions云端访问生产文件。
认证信息来自GitHub Secrets：
- GOOGLE_SERVICE_ACCOUNT_JSON
- DRIVE_ROOT_FOLDER_ID
"""

import json
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]


class DriveManager:
    def __init__(self):
        self.folder_id = os.getenv("DRIVE_ROOT_FOLDER_ID", "")
        self.credentials_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
        self.service = None

    def connect(self):
        if not self.credentials_json:
            raise RuntimeError("缺少 GOOGLE_SERVICE_ACCOUNT_JSON")

        info = json.loads(self.credentials_json)
        credentials = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES
        )
        self.service = build("drive", "v3", credentials=credentials)
        return self.service

    def list_drive_files(self, folder_id=None):
        if self.service is None:
            self.connect()

        folder = folder_id or self.folder_id
        query = f"'{folder}' in parents and trashed=false"
        result = self.service.files().list(
            q=query,
            fields="files(id,name,mimeType,parents)"
        ).execute()
        return result.get("files", [])

    def scan_folder_recursive(self, folder_id=None, path=""):
        """递归扫描Drive目录，返回所有文件信息"""
        files = []
        current_folder = folder_id or self.folder_id

        for item in self.list_drive_files(current_folder):
            item_path = f"{path}/{item['name']}"
            if item.get("mimeType") == "application/vnd.google-apps.folder":
                files.extend(
                    self.scan_folder_recursive(item["id"], item_path)
                )
            else:
                item["path"] = item_path
                files.append(item)

        return files

    def upload_file(self, file_path, target_folder_id=None):
        if self.service is None:
            self.connect()

        metadata = {"name": os.path.basename(file_path)}
        if target_folder_id:
            metadata["parents"] = [target_folder_id]

        media = MediaFileUpload(file_path)
        return self.service.files().create(
            body=metadata,
            media_body=media,
            fields="id"
        ).execute()

    def move_file(self, file_id, target_folder_id):
        if self.service is None:
            self.connect()

        file_info = self.service.files().get(
            fileId=file_id,
            fields="parents"
        ).execute()

        previous = ",".join(file_info.get("parents", []))

        return self.service.files().update(
            fileId=file_id,
            addParents=target_folder_id,
            removeParents=previous,
            fields="id,parents"
        ).execute()
