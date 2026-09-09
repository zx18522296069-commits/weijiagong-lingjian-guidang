"""
Google Drive认证模块

用于GitHub Actions无人值守执行。
凭证从环境变量读取，不保存私钥。

需要配置：
GOOGLE_SERVICE_ACCOUNT_JSON
"""

import json
import os


def get_credentials_info():
    value = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not value:
        raise RuntimeError("缺少 GOOGLE_SERVICE_ACCOUNT_JSON")
    return json.loads(value)


def check_auth_config():
    return bool(os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON"))
