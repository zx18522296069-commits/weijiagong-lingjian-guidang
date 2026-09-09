"""
未加工零件处理核心逻辑

正式规则：
- 按板材编号累计
- 编号末尾“废”保留并区分
- 防止重复录入
- 多日累计已加工数量
- 整单完成后再移除
"""

from collections import defaultdict


def normalize_material_id(value):
    """保留原始编号，包含废字，不做模糊合并"""
    if value is None:
        return ""
    return str(value).strip()


def build_processed_index(records):
    """按板材编号+图号累计已加工数量"""
    result = defaultdict(int)
    for row in records:
        material_id = normalize_material_id(row.get("板材编号"))
        drawing = str(row.get("图号", "")).strip()
        quantity = int(row.get("完成数量", 0) or 0)
        key = (material_id, drawing)
        result[key] += quantity
    return dict(result)


def calculate_remaining(total_count, processed_count):
    """当前剩余 = 应加工数量 - 累计已加工数量"""
    return int(total_count or 0) - int(processed_count or 0)


def get_status(remaining):
    """状态标识：0完成，正数未完成，负数异常"""
    if remaining == 0:
        return "完成"
    if remaining < 0:
        return "异常-超加工"
    return "进行中"


def build_remaining_records(order_records, processed_index):
    """根据订单汇总和累计台账生成当前剩余明细"""
    result = []

    for row in order_records:
        material_id = normalize_material_id(row.get("板材编号"))
        drawing = str(row.get("图号", "")).strip()
        total = int(row.get("件数", 0) or 0)

        processed = processed_index.get((material_id, drawing), 0)
        remaining = calculate_remaining(total, processed)

        result.append({
            **row,
            "累计已加工": processed,
            "当前剩余": remaining,
            "状态": get_status(remaining)
        })

    return result
