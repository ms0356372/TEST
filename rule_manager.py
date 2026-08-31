"""規則驗證與可攜式 JSON 匯入／匯出。"""
from __future__ import annotations

import json
from pathlib import Path


def validate_rules(rules: object) -> list[dict[str, str]]:
    """驗證規則陣列，並保留其排序；來源值不可重複或空白。"""
    if not isinstance(rules, list):
        raise ValueError("rules 必須是陣列")
    result: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in rules:
        if not isinstance(item, dict):
            raise ValueError("每筆規則必須是物件")
        source, target = str(item.get("source", "")).strip(), str(item.get("target", "")).strip()
        if not source or not target:
            raise ValueError("原始值與轉換值不可空白")
        if source in seen:
            raise ValueError(f"規則 {source} 重複")
        seen.add(source)
        result.append({"source": source, "target": target})
    return result


def import_rules(path: Path) -> list[dict[str, str]]:
    """由 UTF-8 JSON 載入規則；接受規則陣列或含 rules 的物件。"""
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    return validate_rules(data.get("rules") if isinstance(data, dict) else data)


def export_rules(path: Path, rules: list[dict[str, str]]) -> None:
    """輸出具版本資訊的規則檔。"""
    payload = {"config_version": 1, "rules": validate_rules(rules)}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
