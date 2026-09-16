from __future__ import annotations

import unittest
from copy import deepcopy

from safe_main import business_outputs_changed, business_ledger_signature, legacy_block_compat_message


class SafeWriteGuardTests(unittest.TestCase):
    def _ledger(self):
        return {
            "historical_rows": [
                {
                    "order": "D53K-1600A-0911",
                    "drawing": "D53K-1600A.3.2-1-7",
                    "thickness": 100,
                    "bevel": "W",
                    "length": 1000,
                    "width": 500,
                    "quantity": 4,
                    "total_weight_t": 1.2,
                    "board_sources": "#100×2",
                    "processed": 2,
                    "remaining": 2,
                    "pending_weight_t": 0.6,
                    "status": "部分完成",
                },
                {
                    "order": "THP10-8000J-0911",
                    "drawing": "2310",
                    "thickness": 80,
                    "bevel": "W",
                    "length": 273,
                    "width": 273,
                    "quantity": 4,
                    "total_weight_t": 0.147,
                    "board_sources": "废209.4×4",
                    "processed": 4,
                    "remaining": 0,
                    "pending_weight_t": 0,
                    "status": "已完成",
                },
            ],
            "flows": [
                {
                    "板材号": "#100",
                    "内容指纹": "abc",
                    "拆图结果文件": "#100_完成.xlsx",
                    "图号": "D53K-1600A.3.2-1-7",
                    "厚度(mm)": 100,
                    "坡口": "W",
                    "本张板加工数量": 2,
                    "数据性质": "正式入账",
                    "说明": "测试",
                }
            ],
            "board_records": [
                {
                    "数据性质": "正式入账",
                    "板材号": "#100",
                    "内容指纹": "abc",
                    "拆图结果文件": "#100_完成.xlsx",
                    "计入件数": 2,
                    "状态": "已核验并入账",
                    "归档位置": "拆图结果/已录入数量",
                }
            ],
            "anomalies": [],
        }

    def test_blocked_anomaly_does_not_count_as_business_change(self):
        existing = self._ledger()
        candidate = deepcopy(existing)
        candidate["anomalies"].append(
            {
                "板材号": "#2203",
                "异常类型": "阻断入账",
                "说明": "无匹配；未入账、未归档",
            }
        )
        self.assertFalse(business_outputs_changed(existing, candidate))

    def test_status_wording_normalization_requires_one_time_write(self):
        existing = self._ledger()
        existing["historical_rows"].append(
            {
                "order": "ORDER-1",
                "drawing": "A1",
                "thickness": 20,
                "bevel": "",
                "length": 10,
                "width": 10,
                "quantity": 2,
                "total_weight_t": 0.1,
                "board_sources": "",
                "processed": 0,
                "remaining": 2,
                "pending_weight_t": 0.1,
                "status": "未开始",
            }
        )
        candidate = deepcopy(existing)
        candidate["historical_rows"][-1]["status"] = "未加工"
        self.assertNotEqual(
            business_ledger_signature(existing),
            business_ledger_signature(candidate),
        )
        self.assertTrue(business_outputs_changed(existing, candidate))

    def test_processed_quantity_change_requires_write(self):
        existing = self._ledger()
        candidate = deepcopy(existing)
        candidate["historical_rows"][0]["processed"] = 3
        candidate["historical_rows"][0]["remaining"] = 1
        self.assertTrue(business_outputs_changed(existing, candidate))

    def test_source_fact_change_requires_write(self):
        existing = self._ledger()
        candidate = deepcopy(existing)
        candidate["historical_rows"][0]["quantity"] = 5
        candidate["historical_rows"][0]["remaining"] = 3
        self.assertTrue(business_outputs_changed(existing, candidate))

    def test_new_flow_or_board_record_requires_write(self):
        existing = self._ledger()
        candidate = deepcopy(existing)
        candidate["flows"].append({"板材号": "#101", "本张板加工数量": 1})
        self.assertTrue(business_outputs_changed(existing, candidate))

        candidate = deepcopy(existing)
        candidate["board_records"].append({"板材号": "#101", "状态": "已核验并入账"})
        self.assertTrue(business_outputs_changed(existing, candidate))

    def test_structured_block_adds_legacy_log_only_when_missing(self):
        seen: set[str] = set()
        self.assertIsNone(legacy_block_compat_message("板材 #2203 阻断：无匹配", seen))
        self.assertIn("#2203", seen)

        message = (
            "板材处理结果｜文件=#2326_完成｜板材=#2326｜状态=未累计、未记录｜"
            "原因=检测到历史入账记录；本次无法复核内容，不重复扣减、不归档。｜"
            "处理建议=核对后重新执行。"
        )
        compat = legacy_block_compat_message(message, seen)
        self.assertEqual(
            compat,
            "板材 #2326 阻断：检测到历史入账记录；本次无法复核内容，不重复扣减、不归档。",
        )
        self.assertIsNone(legacy_block_compat_message(message, seen))


if __name__ == "__main__":
    unittest.main()
