"""
Excel输出模块

正式规则：
- 当前待加工零件
- 累计加工台账
- 全部文字居中
- J列当前剩余重点显示
- 完成状态颜色标识
"""

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


def format_remaining_column(ws, column="J"):
    """当前剩余列格式"""
    for cell in ws[column]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(size=14, bold=True)


def center_all_cells(ws):
    """全部单元格居中"""
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")


def generate_report(data, output_path):
    """
    根据模板生成正式Excel

    后续接入：
    - 模板读取
    - 数据写入
    - 状态颜色
    - 统计页
    """
    pass
