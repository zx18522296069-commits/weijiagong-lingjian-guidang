"""未加工零件安全入口。

在现有稳定业务主流程外增加一层“正式 Excel 写回保护”：
- 阻断板只记录运行日志，不因为新增阻断异常/运行说明而改写正式 Excel；
- 只有订单事实、累计数量、加工流水或板材入账记录真正变化时，才允许覆盖两份正式表；
- 补归档场景可在不重写正式表的情况下继续完成归档；
- 不改变 main.py 现有的匹配、整板校验、幂等和归档规则。
"""

from __future__ import annotations

import json
import math
import os
from copy import deepcopy
from pathlib import Path

from modules.formal_files import CUMULATIVE_NAME, PENDING_NAME
from modules.logger import get_logger
from self_healing_main import prepare_official_files

logger = get_logger()


_ROW_FIELDS = (
    "order",
    "drawing",
    "thickness",
    "bevel",
    "length",
    "width",
    "quantity",
    "total_weight_t",
    "board_sources",
    "processed",
    "remaining",
    "pending_weight_t",
    "status",
)


def _scalar(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        number = float(value)
        if math.isfinite(number):
            return round(number, 9)
        return str(value)
    text = str(value).strip()
    # “未开始”是旧版文案，“未加工”是新版文案；二者业务含义相同，
    # 不能仅因为文案变化就触发正式台账重写。
    if text == "未开始":
        return "未加工"
    return text


def _normalized_dict(row: dict, fields: tuple[str, ...] | None = None) -> dict:
    keys = fields if fields is not None else tuple(sorted(str(key) for key in row.keys()))
    return {key: _scalar(row.get(key)) for key in keys}


def _sorted_records(rows: list[dict], fields: tuple[str, ...] | None = None) -> list[dict]:
    normalized = [_normalized_dict(row, fields) for row in rows]
    return sorted(
        normalized,
        key=lambda item: json.dumps(item, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
    )


def business_ledger_signature(ledger: dict) -> str:
    """只比较正式业务事实；故意忽略“异常记录”和运行说明。

    阻断文件没有正式入账，因此其异常只属于本次运行诊断，不能成为
    覆盖《累计加工台账》《当前待加工零件》的理由。
    """
    payload = {
        "rows": _sorted_records(list(ledger.get("historical_rows", [])), _ROW_FIELDS),
        "flows": _sorted_records(list(ledger.get("flows", []))),
        "board_records": _sorted_records(list(ledger.get("board_records", []))),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def business_outputs_changed(existing: dict, candidate: dict) -> bool:
    return business_ledger_signature(existing) != business_ledger_signature(candidate)


def run() -> None:
    # 先保留现有自愈逻辑：正式表缺失时仍按现有规则安全重建。
    from modules.drive_manager import DriveManager

    drive = DriveManager()
    if not drive.check_config():
        raise RuntimeError("Google Drive 配置不完整")

    official = prepare_official_files(drive)
    missing = [name for name, item in official.items() if item is None]
    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    if missing:
        if test_mode:
            return
        raise RuntimeError(f"正式台账自愈失败：{missing}")

    os.environ["CUMULATIVE_FILE_ID"] = str(official[CUMULATIVE_NAME]["id"])
    os.environ["PENDING_FILE_ID"] = str(official[PENDING_NAME]["id"])

    import main as business_main

    original_read_existing = business_main.read_existing_ledger
    original_write_formal = business_main.write_formal_file
    context: dict[str, object] = {
        "existing": None,
        "business_changed": None,
    }

    def capture_existing(path):
        ledger = original_read_existing(path)
        # main.py 第一次读取的是运行前正式累计台账；后续回读验证不能覆盖快照。
        if context["existing"] is None:
            context["existing"] = deepcopy(ledger)
        return ledger

    def guarded_write_formal(drive_obj, current, local_path):
        name = Path(local_path).name

        if name == CUMULATIVE_NAME:
            if current is None or context["existing"] is None:
                context["business_changed"] = True
                return original_write_formal(drive_obj, current, local_path)

            candidate = original_read_existing(local_path)
            changed = business_outputs_changed(context["existing"], candidate)
            context["business_changed"] = changed
            if not changed:
                logger.info(
                    "写回保护｜累计业务事实无变化：跳过累计加工台账上传；"
                    "阻断异常仅保留运行日志，不改正式台账"
                )
                return {"id": current["id"]}
            logger.info("写回保护｜检测到累计业务事实变化：允许更新累计加工台账")
            return original_write_formal(drive_obj, current, local_path)

        if name == PENDING_NAME:
            if current is not None and context["business_changed"] is False:
                logger.info("写回保护｜累计业务事实无变化：跳过当前待加工零件上传")
                return {"id": current["id"]}
            return original_write_formal(drive_obj, current, local_path)

        return original_write_formal(drive_obj, current, local_path)

    business_main.read_existing_ledger = capture_existing
    business_main.write_formal_file = guarded_write_formal
    business_main.run()


if __name__ == "__main__":
    run()
