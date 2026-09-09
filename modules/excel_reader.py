"""正式 Excel 读取与结构化解析。原始文件只读，不修改。"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook


def _text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _number(value, field: str, *, allow_blank: bool = False) -> float | None:
    if value in (None, ""):
        if allow_blank:
            return None
        raise ValueError(f"字段 {field} 为空")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"字段 {field} 不是有效数字: {value!r}") from exc


def _int_number(value, field: str) -> int:
    num = _number(value, field)
    rounded = int(round(num))
    if abs(num - rounded) > 1e-9:
        raise ValueError(f"字段 {field} 必须为整数: {value!r}")
    return rounded


def normalize_order_key(value: str) -> str:
    """
    将“175.26-08-05  D53K-1600A-0805”确定性归一为 D53K-1600A-0805。
    如果没有前置内部编号，则保留原值。
    """
    text = _text(value)
    if not text:
        return ""
    parts = [p for p in re.split(r"\s+", text) if p]
    if len(parts) >= 2:
        return parts[-1]
    return text


def normalize_drawing(value: str) -> str:
    return _text(value)


def normalize_bevel(value: str) -> str:
    return _text(value)


def _header_map(row: Iterable) -> dict[str, int]:
    result = {}
    for idx, value in enumerate(row, start=1):
        text = _text(value)
        if text:
            result[text] = idx
    return result


def _find_header(ws, required: set[str], max_rows: int = 40):
    for r in range(1, min(ws.max_row, max_rows) + 1):
        values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        texts = {_text(v) for v in values if _text(v)}
        if required.issubset(texts):
            return r, _header_map(values)
    return None, None


def _pick_header(headers: dict[str, int], *candidates: str) -> tuple[str, int] | tuple[None, None]:
    for candidate in candidates:
        if candidate in headers:
            return candidate, headers[candidate]
    return None, None


def read_source_summary(path, *, source_name: str = "", order_container_name: str = "") -> list[dict]:
    """读取订单原始汇总表，重量统一为吨；长宽、数量、坡口直接来自原始汇总表。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    records: list[dict] = []

    for ws in wb.worksheets:
        header_row, headers = _find_header(ws, {"订单号", "图号", "厚度", "件数"})
        if not header_row:
            continue

        _, order_col = _pick_header(headers, "订单号")
        _, drawing_col = _pick_header(headers, "图号")
        _, thickness_col = _pick_header(headers, "厚度", "厚度(mm)")
        _, qty_col = _pick_header(headers, "件数", "数量", "订单总数量")
        _, length_col = _pick_header(headers, "长(mm)", "长")
        _, width_col = _pick_header(headers, "宽(mm)", "宽")
        _, bevel_col = _pick_header(headers, "坡口")
        weight_name, weight_col = _pick_header(
            headers,
            "总净重(t)", "总净重（t）", "总净重(T)", "总净重（T）",
            "总净重(kg)", "总净重（kg）", "总净重(KG)", "总净重（KG）",
        )
        if not all([order_col, drawing_col, thickness_col, qty_col, length_col, width_col, bevel_col, weight_col]):
            raise ValueError(f"原始汇总表缺少正式必需字段: {source_name or path}")

        weight_is_kg = "kg" in weight_name.lower()

        for r in range(header_row + 1, ws.max_row + 1):
            order_raw = _text(ws.cell(r, order_col).value)
            drawing = normalize_drawing(ws.cell(r, drawing_col).value)
            if not order_raw and not drawing:
                continue
            if not order_raw:
                # 合计、说明、厚度核对等尾部区域不属于零件明细。
                continue
            if not drawing:
                raise ValueError(f"原始汇总表第{r}行图号为空: {source_name or path}")

            total_weight = _number(ws.cell(r, weight_col).value, "总净重")
            total_weight_t = total_weight / 1000.0 if weight_is_kg else total_weight
            records.append({
                "order_raw": order_raw,
                "order": normalize_order_key(order_raw),
                "drawing": drawing,
                "thickness": _number(ws.cell(r, thickness_col).value, "厚度"),
                "quantity": _int_number(ws.cell(r, qty_col).value, "件数"),
                "length": _number(ws.cell(r, length_col).value, "长(mm)"),
                "width": _number(ws.cell(r, width_col).value, "宽(mm)"),
                "bevel": normalize_bevel(ws.cell(r, bevel_col).value),
                "total_weight_t": float(total_weight_t),
                "source_file": source_name or Path(path).name,
                "order_container_name": order_container_name,
                "source_row": r,
            })

        if records:
            return records

    raise ValueError(f"未找到订单汇总表正式表头: {source_name or path}")


def read_split_result(path, *, board_id: str, source_name: str = "") -> dict:
    """读取 `编号_完成.xlsx` 结构化明细。坡口若不存在则保留为空，由匹配阶段做唯一性检查。"""
    wb = load_workbook(path, read_only=True, data_only=True)

    for ws in wb.worksheets:
        header_row, headers = _find_header(ws, {"订单号", "图号", "厚度"})
        if not header_row:
            continue

        qty_name, split_qty_col = _pick_header(headers, "图片拆分数量", "拆分数量", "本次加工数量")
        base_qty_name, base_qty_col = _pick_header(headers, "基础件数", "基础数量")
        base_weight_name, base_weight_col = _pick_header(
            headers,
            "基础表总重量（T）", "基础表总重量(T)", "基础表总重量（t）", "基础表总重量(t)",
        )
        _, order_col = _pick_header(headers, "订单号")
        _, drawing_col = _pick_header(headers, "图号")
        _, thickness_col = _pick_header(headers, "厚度", "厚度(mm)")
        _, bevel_col = _pick_header(headers, "坡口")

        if not all([split_qty_col, base_qty_col, base_weight_col, order_col, drawing_col, thickness_col]):
            raise ValueError(f"拆图结果缺少正式必需字段: {source_name or path}")

        rows: list[dict] = []
        for r in range(header_row + 1, ws.max_row + 1):
            order_raw = _text(ws.cell(r, order_col).value)
            drawing = normalize_drawing(ws.cell(r, drawing_col).value)
            if not order_raw and not drawing:
                continue
            if not order_raw or not drawing:
                # 后部复核说明区不属于结构化明细。
                continue
            split_qty = _int_number(ws.cell(r, split_qty_col).value, qty_name)
            if split_qty <= 0:
                raise ValueError(f"拆图结果加工数量必须>0: {source_name or path} 第{r}行")
            rows.append({
                "board_id": board_id,
                "source_file": source_name or Path(path).name,
                "order_raw": order_raw,
                "order": normalize_order_key(order_raw),
                "drawing": drawing,
                "thickness": _number(ws.cell(r, thickness_col).value, "厚度"),
                "bevel": normalize_bevel(ws.cell(r, bevel_col).value) if bevel_col else "",
                "base_quantity": _int_number(ws.cell(r, base_qty_col).value, base_qty_name),
                "split_quantity": split_qty,
                "base_total_weight_t": _number(ws.cell(r, base_weight_col).value, base_weight_name),
                "source_row": r,
            })

        review = {}
        labels = {
            "图片标注重量（KG）": "image_weight_kg",
            "图片标注重量(KG)": "image_weight_kg",
            "拆分重量合计（KG）": "split_weight_kg",
            "拆分重量合计(KG)": "split_weight_kg",
            "差额绝对值（KG）": "abs_diff_kg",
            "差额绝对值(KG)": "abs_diff_kg",
            "复核结论": "review_conclusion",
        }
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                label = _text(ws.cell(r, c).value)
                if label not in labels:
                    continue
                value = None
                for cc in range(c + 1, min(ws.max_column, c + 3) + 1):
                    candidate = ws.cell(r, cc).value
                    if candidate not in (None, ""):
                        value = candidate
                        break
                if value not in (None, ""):
                    review[labels[label]] = value

        if rows:
            return {"rows": rows, "review": review}

    raise ValueError(f"未找到拆图结果正式表头或无结构化明细: {source_name or path}")


def read_existing_ledger(path) -> dict:
    """读取正式累计台账中的当前累计、加工流水、板材入账和异常历史。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    if "累计加工台账" not in wb.sheetnames:
        raise ValueError("累计加工台账.xlsx 缺少“累计加工台账”工作表")

    state: dict[tuple, dict] = {}
    ws = wb["累计加工台账"]
    current_order = ""
    header_cols = None
    for r in range(1, ws.max_row + 1):
        first = _text(ws.cell(r, 1).value)
        if first.startswith("订单："):
            text = first[len("订单："):]
            current_order = normalize_order_key(text.split("｜", 1)[0].strip())
            header_cols = None
            continue
        if first == "图号":
            headers = _header_map([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
            header_cols = headers
            continue
        if not current_order or not header_cols or not first:
            continue

        drawing = normalize_drawing(ws.cell(r, header_cols.get("图号", 1)).value)
        if not drawing:
            continue
        thickness_col = header_cols.get("厚度(mm)") or header_cols.get("厚度")
        bevel_col = header_cols.get("坡口")
        processed_col = header_cols.get("累计已加工")
        source_col = header_cols.get("板材号/加工来源")
        if not all([thickness_col, bevel_col, processed_col]):
            raise ValueError("累计加工台账主表字段结构不符合正式模板")
        key = (
            current_order,
            drawing,
            float(_number(ws.cell(r, thickness_col).value, "厚度")),
            normalize_bevel(ws.cell(r, bevel_col).value),
        )
        state[key] = {
            "processed": _int_number(ws.cell(r, processed_col).value or 0, "累计已加工"),
            "board_sources": _text(ws.cell(r, source_col).value) if source_col else "",
        }

    def read_sheet_records(sheet_name: str) -> list[dict]:
        if sheet_name not in wb.sheetnames:
            return []
        sheet = wb[sheet_name]
        if sheet.max_row < 1:
            return []
        headers = [_text(sheet.cell(1, c).value) for c in range(1, sheet.max_column + 1)]
        records = []
        for r in range(2, sheet.max_row + 1):
            values = [sheet.cell(r, c).value for c in range(1, sheet.max_column + 1)]
            if not any(v not in (None, "") for v in values):
                continue
            records.append({headers[i]: values[i] for i in range(len(headers)) if headers[i]})
        return records

    flows = read_sheet_records("加工流水")
    board_records = read_sheet_records("板材入账记录")
    anomalies = read_sheet_records("异常记录")
    posted_boards = {
        _text(row.get("板材号"))
        for row in board_records
        if _text(row.get("板材号"))
        and _text(row.get("状态")) not in {"作废", "未入账", "阻断"}
    }

    return {
        "state": state,
        "flows": flows,
        "board_records": board_records,
        "anomalies": anomalies,
        "posted_boards": posted_boards,
    }


def read_excel(path):
    """兼容旧接口：返回第一张表的二维值，不参与正式业务逻辑。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    return [list(row) for row in ws.iter_rows(values_only=True)]
