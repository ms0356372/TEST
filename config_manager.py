"""使用者設定的建立、修復、載入與儲存。"""
from __future__ import annotations

import json
import os
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from rule_manager import validate_rules

DEFAULT_CONFIG = {
    "config_version": 1, "start_cell": "E4", "input_delimiter": ";",
    "output_delimiter": ";", "remove_duplicates": True,
    "unknown_behavior": "keep", "result_header": "特殊作業轉換",
    "create_header": True, "last_sheet": "",
    "rules": [{"source": "99", "target": "一般"}, {"source": "02", "target": "噪音"},
              {"source": "03", "target": "游離輻射"}],
}


def get_app_dir() -> Path:
    """Windows 使用 APPDATA；其他平台使用相容的使用者設定目錄。"""
    base = Path(os.environ.get("APPDATA", Path.home() / ".config"))
    return base / "ExcelTransformTool"


def get_config_path() -> Path:
    return get_app_dir() / "config.json"


def save_config(config: dict, path: Path | None = None) -> Path:
    """先驗證再以 replace 原子式寫入，避免中途關閉造成損壞。"""
    target = path or get_config_path(); target.parent.mkdir(parents=True, exist_ok=True)
    merged = deepcopy(DEFAULT_CONFIG); merged.update(config); merged["rules"] = validate_rules(merged["rules"])
    if merged["unknown_behavior"] not in ("keep", "ignore"):
        raise ValueError("未知代碼處理方式無效")
    temp = target.with_suffix(".tmp")
    temp.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(target)
    return target


def load_config(path: Path | None = None) -> tuple[dict, bool]:
    """讀取設定；損壞時備份並重建，第二個回傳值代表曾修復。"""
    target = path or get_config_path(); target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists():
        save_config(DEFAULT_CONFIG, target); return deepcopy(DEFAULT_CONFIG), False
    try:
        raw = json.loads(target.read_text(encoding="utf-8-sig"))
        if not isinstance(raw, dict): raise ValueError("設定根節點必須是物件")
        merged = deepcopy(DEFAULT_CONFIG); merged.update(raw); merged["rules"] = validate_rules(merged["rules"])
        return merged, False
    except (OSError, ValueError, json.JSONDecodeError, TypeError):
        backup = target.with_name(f"config_corrupted_{datetime.now():%Y%m%d_%H%M%S}.json")
        try: target.replace(backup)
        except OSError: pass
        save_config(DEFAULT_CONFIG, target)
        return deepcopy(DEFAULT_CONFIG), True
