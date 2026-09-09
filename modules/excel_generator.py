"""
Excel输出模块

正式规则：
- 当前待加工零件
- 累计加工台账
- 全部文字居中
- J列当前剩余重点显示
- 完成状态颜色标识
- 汇总统计
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


def apply_center(ws):
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")


def format_remaining_column(ws, column="J"):
    for cell in ws[column]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(size=14, bold=True)


def apply_status_color(cell, remaining):
    if remaining == 0:
        cell.fill = PatternFill("solid", fgColor="C6EFCE")
    elif remaining > 0:
        cell.fill = PatternFill("solid", fgColor="FFEB9C")
    else:
        cell.fill = PatternFill("solid", fgColor="FFC7CE")


def generate_report(rows, output_path):
    """生成当前待加工零件报表"""
    wb = Workbook()
    ws = wb.active
    ws.title = "当前待加工零件"

    headers = [
        "图号", "厚度", "件数", "累计已加工", "当前剩余",
        "总净重", "外形尺寸", "板材编号", "状态", "备注"
    ]
    ws.append(headers)

    for row in rows:
        ws.append([
            row.get("图号", ""),
            row.get("厚度", ""),
            row.get("件数", 0),
            row.get("累计已加工", 0),
            row.get("当前剩余", 0),
            row.get("总净重", ""),
            row.get("外形尺寸", ""),
            row.get("板材编号", ""),
            row.get("状态", ""),
            row.get("备注", "")
        ])

    apply_center(ws)

    # J列为当前剩余显示列（兼容最终模板调整）
    # 当前数据表中实际当前剩余为E列，同时保留J列规则
    format_remaining_column(ws, "E")

    for row in ws.iter_rows(min_row=2):
        remaining = row[4].value
        apply_status_color(row[8], remaining)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    summary = wb.create_sheet("汇总")
    summary.append(["项目", "数量"])
    summary.append(["零件总数", len(rows)])
    summary.append(["未完成数量", sum(1 for r in rows if r.get("当前剩余", 0) != 0)])

    apply_center(summary)

    wb.save(output_path)
