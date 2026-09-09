"""
Google Drive连接模块

用于GitHub Actions云端访问：
- 正在加工
- 拆图结果
- 输出文件

实际凭证通过GitHub Secrets注入，不保存账号信息。
"""


def list_drive_files(folder_id):
    """读取指定Drive文件夹，待接入Google Drive API"""
    raise NotImplementedError


def upload_file(file_path, target_folder_id):
    """上传生成文件，待接入Google Drive API"""
    raise NotImplementedError
