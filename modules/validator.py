"""
数据核对规则

用于检查：
- 重复录入
- 数量异常
- 编号冲突
- 废字编号区分
"""


def check_duplicate(material_id, imported_ids):
    return material_id in imported_ids


def check_remaining(value):
    return value < 0
