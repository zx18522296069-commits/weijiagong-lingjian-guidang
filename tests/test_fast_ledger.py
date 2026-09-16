from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from modules.excel_generator import generate_cumulative_report
from modules.excel_reader import read_existing_ledger
from modules.fast_ledger import read_existing_ledger_fast


class FastLedgerEquivalenceTests(unittest.TestCase):
    def test_fast_reader_matches_existing_reader(self):
        rows = [
            {
                "order_raw": "ORDER-1",
                "order": "ORDER-1",
                "drawing": "A-01",
                "thickness": 40.0,
                "bevel": "W",
                "length": 1000.0,
                "width": 500.0,
                "quantity": 4,
                "total_weight_t": 1.2,
                "source_file": "source.xlsx",
                "order_container_name": "ORDER-1",
                "source_row": 2,
                "board_sources": "#100×2",
                "processed": 2,
                "remaining": 2,
                "pending_weight_t": 0.6,
                "status": "部分完成",
            }
        ]
        flows = [{
            "板材号": "#100", "内容指纹": "abc", "拆图结果文件": "#100_完成.xlsx",
            "图号": "A-01", "厚度(mm)": 40, "坡口": "W", "本张板加工数量": 2,
            "数据性质": "正式入账", "说明": "test",
        }]
        board_records = [{
            "数据性质": "正式入账", "板材号": "#100", "内容指纹": "abc",
            "拆图结果文件": "#100_完成.xlsx", "计入件数": 2,
            "状态": "已核验并入账", "归档位置": "拆图结果/已录入数量",
        }]
        anomalies = [{"板材号": "#999", "异常类型": "阻断入账", "说明": "test"}]

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "累计加工台账.xlsx"
            generate_cumulative_report(rows, flows, board_records, anomalies, path, note="equivalence")
            old = read_existing_ledger(path)
            new = read_existing_ledger_fast(path)

        self.assertEqual(old, new)


if __name__ == "__main__":
    unittest.main()
