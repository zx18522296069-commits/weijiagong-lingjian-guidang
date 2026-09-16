"""临时验收工具：完整计算候选台账，但禁止正式写回和 runtime state 覆盖。"""
from __future__ import annotations

import json

import safe_main
from modules.runtime_state import RuntimeState
from production_main import install_drive_guards


def _canon(rows, fields=None):
    normalized = safe_main._sorted_records(list(rows or []), fields)
    return {
        json.dumps(item, ensure_ascii=False, sort_keys=True, separators=(",", ":")): item
        for item in normalized
    }


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
        old = _canon(old_rows, fields)
        new = _canon(new_rows, fields)
        removed_keys = sorted(set(old) - set(new))
        added_keys = sorted(set(new) - set(old))
        if removed_keys or added_keys:
            changed = True
            print(f"WRITE_GUARD_DIFF section={name} old={len(old)} new={len(new)} removed={len(removed_keys)} added={len(added_keys)}")
            for key in removed_keys[:3]:
                print("WRITE_GUARD_REMOVED", name, json.dumps(old[key], ensure_ascii=False, sort_keys=True))
            for key in added_keys[:3]:
                print("WRITE_GUARD_ADDED", name, json.dumps(new[key], ensure_ascii=False, sort_keys=True))
        else:
            print(f"WRITE_GUARD_SAME section={name} count={len(old)}")
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
