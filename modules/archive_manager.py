"""
已录入数量归档模块

正式规则：
- 单张板材完成文件经核对、写入永久台账并回读验证后即可归档
- 归档只负责避免完成文件重复扫描，不删除永久历史
- “整单完成才移除”只适用于《当前待加工零件》，不限制完成文件归档
"""

from datetime import datetime


def should_archive(posted_and_verified):
    """只有该板材已正式入账且回读验证通过，才允许归档完成文件。"""
    return posted_and_verified is True


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
