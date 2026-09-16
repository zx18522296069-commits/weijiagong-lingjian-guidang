"""“正在加工”根目录正式台账文件发现规则。"""

from __future__ import annotations

from modules.drive_manager import FOLDER_MIME, SHORTCUT_MIME

CUMULATIVE_NAME = "累计加工台账.xlsx"
PENDING_NAME = "当前待加工零件.xlsx"
OFFICIAL_NAMES = (CUMULATIVE_NAME, PENDING_NAME)


def discover_official_files(items: list[dict]) -> dict[str, dict | None]:
    """
    按正式文件名从“正在加工”根目录直接子项中定位两个正式 Excel。

    - 不依赖历史固定 file ID；删除后重建导致 ID 改变也能继续工作。
    - 同名真实文件多于一个时必须阻断，禁止猜测使用哪一个。
    - 文件夹、快捷方式或近似名称都不能冒充正式表。
    """
    result: dict[str, dict | None] = {}
    for name in OFFICIAL_NAMES:
        matches = [
            item for item in items
            if item.get("mimeType") not in {FOLDER_MIME, SHORTCUT_MIME}
            and str(item.get("name", "")) == name
        ]
        if len(matches) > 1:
            ids = [str(item.get("id", "")) for item in matches]
            raise RuntimeError(f"正在加工根目录存在多个同名正式文件，禁止自动猜测：{name} -> {ids}")
        result[name] = matches[0] if matches else None
    return result
