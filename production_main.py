"""未加工零件正式生产入口。

在 safe_main 外增加 Google Drive 网络可靠性与同轮 metadata 快照层：
- list/get/download 遇到临时 SSL、网络、429、5xx 时有限重试；
- 同一轮扫描得到的 metadata 复用，避免每个正文下载前再次请求 Drive；
- 正式 Excel 更新后立即使旧 metadata 失效，上传后的独立校验 GET 必须真实访问 Drive；
- 校验 GET 成功后，runtime state 提交复用这份已经验证过的 metadata，不再额外请求一次 Drive；
- 下载完成后按扫描时 md5Checksum 核对正文，发现扫描后文件变化则重新获取 metadata 后再试。

业务规则仍全部由 safe_main.py / main.py 执行，本文件不改变任何零件匹配、累计、归档规则。
"""

from __future__ import annotations

import hashlib
import ssl
import time
from pathlib import Path

from googleapiclient.errors import HttpError

from modules.drive_manager import DriveManager
from modules.logger import get_logger

logger = get_logger()


class SnapshotChangedError(RuntimeError):
    pass


def _retryable(error: Exception) -> bool:
    if isinstance(error, (ssl.SSLError, OSError, ConnectionError, TimeoutError, SnapshotChangedError)):
        return True
    if isinstance(error, HttpError):
        status = int(getattr(error.resp, "status", 0) or 0)
        return status in {408, 409, 429} or status >= 500
    return False


def _retry(label: str, func, *, attempts: int = 4):
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return func()
        except Exception as error:
            last_error = error
            if attempt >= attempts or not _retryable(error):
                raise
            delay = min(2 ** (attempt - 1), 4)
            logger.warning(f"Drive临时异常，准备重试｜操作={label}｜第{attempt}/{attempts}次｜{type(error).__name__}: {error}")
            time.sleep(delay)
    raise last_error  # pragma: no cover


def _md5(path: Path) -> str:
    digest = hashlib.md5()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def install_drive_guards() -> None:
    if getattr(DriveManager, "_production_guards_installed", False):
        return

    original_list_children = DriveManager.list_children
    original_get_file = DriveManager.get_file
    original_download_file = DriveManager.download_file
    original_update_file_content = DriveManager.update_file_content

    metadata_cache: dict[str, dict] = {}
    force_verify: set[str] = set()

    def remember(items):
        for item in items or []:
            file_id = str(item.get("id", ""))
            if file_id:
                metadata_cache[file_id] = dict(item)

    def requested_fields_available(metadata: dict, fields: str) -> bool:
        requested = [part.strip() for part in str(fields or "").split(",") if part.strip()]
        return bool(requested) and all(field in metadata for field in requested)

    def guarded_list_children(self, folder_id: str):
        items = _retry(f"list_children:{folder_id}", lambda: original_list_children(self, folder_id))
        remember(items)
        return items

    def guarded_get_file(self, file_id: str, fields: str = "id,name,mimeType,parents,modifiedTime,md5Checksum,size"):
        key = str(file_id)
        cached = metadata_cache.get(key)
        if key not in force_verify and cached is not None and requested_fields_available(cached, fields):
            return dict(cached)

        result = _retry(
            f"get_file:{key}",
            lambda: original_get_file(self, file_id, fields=fields),
        )
        metadata_cache[key] = dict(result)
        force_verify.discard(key)
        return result

    def guarded_download_file(self, file_id: str, save_path: str):
        key = str(file_id)
        target = Path(save_path)

        def attempt_download():
            try:
                result = original_download_file(self, file_id, save_path)
                metadata = metadata_cache.get(key) or {}
                expected_md5 = str(metadata.get("md5Checksum") or "").lower()
                if expected_md5 and target.exists():
                    actual_md5 = _md5(target)
                    if actual_md5 != expected_md5:
                        metadata_cache.pop(key, None)
                        target.unlink(missing_ok=True)
                        raise SnapshotChangedError(
                            f"文件在扫描后发生变化：file_id={key}，扫描MD5={expected_md5}，下载MD5={actual_md5}"
                        )
                return result
            except Exception:
                if target.exists():
                    target.unlink(missing_ok=True)
                raise

        return _retry(f"download_file:{key}", attempt_download)

    def guarded_update_file_content(self, file_id: str, file_path: str):
        key = str(file_id)
        metadata_cache.pop(key, None)
        force_verify.add(key)

        def do_update():
            return original_update_file_content(self, file_id, file_path)

        # 覆盖同一个 file_id 且内容相同，临时网络失败后重试不会创建重复正式文件。
        result = _retry(f"update_file_content:{key}", do_update)
        # 故意不把 update 返回值放入 metadata_cache：后续上传校验必须执行一次真实 GET。
        return result

    DriveManager.list_children = guarded_list_children
    DriveManager.get_file = guarded_get_file
    DriveManager.download_file = guarded_download_file
    DriveManager.update_file_content = guarded_update_file_content
    DriveManager._production_guards_installed = True


def run() -> None:
    install_drive_guards()
    from safe_main import run as run_safe

    run_safe()


if __name__ == "__main__":
    run()
