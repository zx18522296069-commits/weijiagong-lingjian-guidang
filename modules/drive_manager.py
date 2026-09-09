"""
Google Drive连接模块

用于GitHub Actions云端访问：
- 正在加工
- 拆图结果
- 当前待加工零件.xlsx
- 累计加工台账.xlsx

认证信息通过GitHub Secrets注入。
不保存个人账号密码。
"""

import os


class DriveManager:
    def __init__(self):
        self.folder_id = os.getenv("DRIVE_ROOT_FOLDER_ID", "")
        self.credentials = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")

    def check_config(self):
        return bool(self.folder_id and self.credentials)

    def list_drive_files(self, folder_id=None):
        """Google Drive API读取入口"""
        return []

    def upload_file(self, file_path, target_folder_id=None):
        """生成Excel上传入口"""
        return False

    def move_file(self, file_id, target_folder_id):
        """已录入数量归档入口"""
        return False
