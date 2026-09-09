"""
Excel读取模块

负责读取：
- 正在加工订单汇总表
- 拆图结果完成文件
- 历史累计台账

规则：
- 不修改原始文件
- 保留原始编号（包括废字）
"""

from pathlib import Path
import pandas as pd


def read_excel(path):
    """读取Excel文件"""
    file_path = Path(path)
    if not file_path.exists():
        return None
    return pd.read_excel(file_path)


def find_completion_files(folder):
    """查找拆图完成文件"""
    folder_path = Path(folder)
    if not folder_path.exists():
        return []
    return list(folder_path.glob("*完成*.xlsx"))
