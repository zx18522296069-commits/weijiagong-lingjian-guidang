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
    """全部单元格居中"""
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")


def format_remaining_column(ws, column="J"):
    """当前剩余列重点显示"""
    for cell in ws[column]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(size=14, bold=True)


def apply_status_color(cell, remaining):
    """0完成，正数进行中，负数异常"""
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
        "总净重", "备注"
    ]
    ws.append(headers)

    for row in rows:
        ws.append(row)

    apply_center(ws)
    format_remaining_column(ws, "E")

    for cell in ws[1]:
        cell.font = Font(bold=True)

    wb.save(output_path)
