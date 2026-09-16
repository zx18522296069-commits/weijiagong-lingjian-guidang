"""临时验收工具：完整计算候选台账，但禁止正式写回和 runtime state 覆盖。"""
from __future__ import annotations

from collections import Counter
import json

import safe_main
from modules.runtime_state import RuntimeState
from production_main import install_drive_guards


def _canon_list(rows, fields=None):
    normalized = safe_main._sorted_records(list(rows or []), fields)
    return [
        json.dumps(item, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        for item in normalized
    ]


def diagnose(existing: dict, candidate: dict) -> bool:
    sections = {
        "rows": (
            existing.get("historical_rows", []),
            candidate.get("historical_rows", []),
            safe_main._ROW_FIELDS,
        ),
        "flows": (existing.get("flows", []), candidate.get("flows", []), None),
        "board_records": (existing.get("board_records", []), candidate.get("board_records", []), None),
        "business_anomalies": (
            [r for r in existing.get("anomalies", []) if str(r.get("异常类型", "")).strip() == "超加工/待核查"],
            [r for r in candidate.get("anomalies", []) if str(r.get("异常类型", "")).strip() == "超加工/待核查"],
            None,
        ),
    }
    changed = False
    for name, (old_rows, new_rows, fields) in sections.items():
        old_list = _canon_list(old_rows, fields)
        new_list = _canon_list(new_rows, fields)
        old = Counter(old_list)
        new = Counter(new_list)
        removed = old - new
        added = new - old
        if removed or added:
            changed = True
            print(
                f"WRITE_GUARD_DIFF section={name} old_raw={len(old_list)} new_raw={len(new_list)} "
                f"old_unique={len(old)} new_unique={len(new)} removed_total={sum(removed.values())} added_total={sum(added.values())}"
            )
            for key, count in list(removed.items())[:5]:
                print("WRITE_GUARD_REMOVED", name, f"count={count}", key)
            for key, count in list(added.items())[:5]:
                print("WRITE_GUARD_ADDED", name, f"count={count}", key)
        else:
            print(f"WRITE_GUARD_SAME section={name} raw={len(old_list)} unique={len(old)}")
    print(f"WRITE_GUARD_DIAGNOSTIC original_changed={changed}; forcing_no_write=true")
    return False


def _noop(*args, **kwargs):
    return None


if __name__ == "__main__":
    install_drive_guards()
    safe_main.business_outputs_changed = diagnose
    RuntimeState.set_cumulative_ledger = _noop
    RuntimeState.capture_source_cache = _noop
    RuntimeState.save = _noop
    safe_main.run()
