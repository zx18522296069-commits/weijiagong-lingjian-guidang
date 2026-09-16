from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from modules.excel_generator import generate_cumulative_report, generate_pending_report
from modules.process_parts import build_current_state, validate_new_board


class BusinessRulesRegressionTests(unittest.TestCase):
    def _source(self):
        return [{
            "order": "ORDER-1",
            "drawing": "A1",
            "thickness": 20.0,
            "bevel": "",
            "length": 1000.0,
            "width": 500.0,
            "quantity": 3,
            "total_weight_t": 0.3,
        }]

    def _split(self, qty=1):
        return {"rows": [{
            "order": "ORDER-1",
            "drawing": "A1",
            "thickness": 20.0,
            "bevel": "",
            "base_quantity": 3,
            "split_quantity": qty,
            "base_total_weight_t": 0.3,
        }]}

    def _color_source(self):
        rows = []
        for drawing, quantity in (
            ("UNTOUCHED", 3),
            ("PARTIAL", 3),
            ("DONE", 2),
            ("OVER", 2),
        ):
            rows.append({
                "order": "COLOR-ORDER",
                "drawing": drawing,
                "thickness": 20.0,
                "bevel": "",
                "length": 1000.0,
                "width": 500.0,
                "quantity": quantity,
                "total_weight_t": quantity * 0.1,
            })
        return rows

    @staticmethod
    def _find_row(ws, drawing: str) -> int:
        return next(
            row for row in range(1, ws.max_row + 1)
            if ws.cell(row, 1).value == drawing
        )

    @staticmethod
    def _assert_row_fill(testcase, ws, row: int, expected: str | None):
        for col in range(1, 13):
            cell = ws.cell(row, col)
            if expected is None:
                testcase.assertIsNone(cell.fill.fill_type)
            else:
                testcase.assertEqual(cell.fill.fill_type, "solid")
                testcase.assertEqual(cell.fill.fgColor.rgb[-6:], expected)

    def test_same_full_board_id_with_different_content_is_blocked(self):
        result = validate_new_board(
            board_id="#2326",
            filename="#2326_完成.xlsx",
            split_payload=self._split(qty=2),
            source_records=self._source(),
            posted_boards={"#2326"},
        )
        self.assertFalse(result["ok"])
        self.assertIn("同板材号内容冲突", result["error"])
        self.assertEqual(result["anomaly"]["异常类型"], "同板材号内容冲突")

    def test_new_board_is_still_allowed(self):
        result = validate_new_board(
            board_id="#2326-1",
            filename="#2326-1_完成.xlsx",
            split_payload=self._split(qty=1),
            source_records=self._source(),
            posted_boards={"#2326"},
        )
        self.assertTrue(result["ok"], result.get("error"))

    def test_all_status_colors_are_identical_in_both_formal_reports(self):
        source = self._color_source()
        existing_state = {
            ("COLOR-ORDER", "PARTIAL", 20.0, ""): {"processed": 1, "board_sources": "#101×1"},
            ("COLOR-ORDER", "DONE", 20.0, ""): {"processed": 2, "board_sources": "#102×2"},
            ("COLOR-ORDER", "OVER", 20.0, ""): {"processed": 3, "board_sources": "#103×3"},
        }
        state = build_current_state(source, existing_state=existing_state)
        statuses = {row["drawing"]: row["status"] for row in state}
        self.assertEqual(statuses, {
            "UNTOUCHED": "未加工",
            "PARTIAL": "部分完成",
            "DONE": "已完成",
            "OVER": "超加工/待核查",
        })

        expected_colors = {
            "UNTOUCHED": None,
            "PARTIAL": "FFF2CC",
            "DONE": "E2F0D9",
            "OVER": "F4CCCC",
        }

        with tempfile.TemporaryDirectory() as td:
            cumulative = Path(td) / "累计加工台账.xlsx"
            pending = Path(td) / "当前待加工零件.xlsx"
            generate_cumulative_report(state, [], [], [], cumulative)
            generate_pending_report(state, pending)

            for path, sheet_name in (
                (cumulative, "累计加工台账"),
                (pending, "当前待加工零件"),
            ):
                ws = load_workbook(path)[sheet_name]
                for drawing, expected in expected_colors.items():
                    row = self._find_row(ws, drawing)
                    self._assert_row_fill(self, ws, row, expected)

    def test_overprocessing_keeps_negative_remaining(self):
        key = ("ORDER-1", "A1", 20.0, "")
        state = build_current_state(
            self._source(),
            existing_state={key: {"processed": 2, "board_sources": "#100×2"}},
            delta_by_key={key: 2},
        )
        self.assertEqual(state[0]["processed"], 4)
        self.assertEqual(state[0]["remaining"], -1)
        self.assertEqual(state[0]["status"], "超加工/待核查")


if __name__ == "__main__":
    unittest.main()
