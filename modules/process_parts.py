"""
未加工零件处理核心逻辑

正式规则：
- 按板材编号累计
- 编号末尾“废”保留并区分
- 防止重复录入
- 多日累计已加工数量
- 整单完成后再移除
"""


def normalize_material_id(value):
    """保留原始编号，包含废字，不做模糊合并"""
    if value is None:
        return ""
    return str(value).strip()


def calculate_remaining(total_count, processed_count):
    return total_count - processed_count
