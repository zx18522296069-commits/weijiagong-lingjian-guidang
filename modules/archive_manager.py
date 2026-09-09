"""
已录入数量归档模块

正式规则：
- 完成核对并入账后归档
- 避免重复扫描
- 保留历史记录
- 整单完成后才允许移出
"""

from datetime import datetime


def should_archive(order_status):
    """只有整单完成才归档"""
    return order_status == "完成"


def build_archive_record(order_id, files=None):
    """生成归档记录，不执行实际移动"""
    return {
        "订单号": order_id,
        "文件数量": len(files or []),
        "归档时间": datetime.now().isoformat(),
        "状态": "待移动"
    }


def archive_completed_file(file_id, target_folder=None):
    """
    移动已完成文件

    实际移动由 Google Drive 模块执行。
    """
    if not file_id:
        return False

    return {
        "file_id": file_id,
        "target_folder": target_folder,
        "archive_time": datetime.now().isoformat()
    }
