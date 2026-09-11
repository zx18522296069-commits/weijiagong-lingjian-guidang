"""Google Drive authentication for unattended GitHub Actions runs.

Personal OAuth is preferred for My Drive because service accounts have no
personal storage quota. Service-account JSON remains supported for Shared
Drive compatibility.
"""

import json
import os

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive"]
TOKEN_URI = "https://oauth2.googleapis.com/token"
OAUTH_ENV_NAMES = (
    "GOOGLE_OAUTH_CLIENT_ID",
    "GOOGLE_OAUTH_CLIENT_SECRET",
    "GOOGLE_OAUTH_REFRESH_TOKEN",
)


def get_credentials_info():
    value = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not value:
        raise RuntimeError("缺少 GOOGLE_SERVICE_ACCOUNT_JSON")
    return json.loads(value)


def check_auth_config():
    oauth_values = [os.getenv(name, "").strip() for name in OAUTH_ENV_NAMES]
    return all(oauth_values) or bool(os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip())


def build_credentials():
    oauth = {name: os.getenv(name, "").strip() for name in OAUTH_ENV_NAMES}
    if any(oauth.values()):
        missing = [name for name, value in oauth.items() if not value]
        if missing:
            raise RuntimeError(f"Google OAuth 配置不完整，缺少: {', '.join(missing)}")
        return Credentials(
            token=None,
            refresh_token=oauth["GOOGLE_OAUTH_REFRESH_TOKEN"],
            token_uri=TOKEN_URI,
            client_id=oauth["GOOGLE_OAUTH_CLIENT_ID"],
            client_secret=oauth["GOOGLE_OAUTH_CLIENT_SECRET"],
            scopes=SCOPES,
        )

    value = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not value:
        raise RuntimeError(
            "缺少 Google 凭据：请配置 GOOGLE_OAUTH_CLIENT_ID、"
            "GOOGLE_OAUTH_CLIENT_SECRET、GOOGLE_OAUTH_REFRESH_TOKEN"
        )
    try:
        info = json.loads(value)
    except json.JSONDecodeError as exc:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON 不是有效 JSON") from exc
    return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
