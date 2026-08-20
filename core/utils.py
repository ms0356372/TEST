from __future__ import annotations
from datetime import datetime
from math import isfinite
from pathlib import Path
from typing import Any

INVALID_NUMBERS = {"", "-", "N/A", "NA", "未檢", "NONE", "NULL"}

def safe_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool): return None
    if isinstance(value, (int, float)):
        return float(value) if isfinite(float(value)) else None
    text = str(value).strip()
    if text.upper() in INVALID_NUMBERS: return None
    try:
        number = float(text.replace(",", ""))
        return number if isfinite(number) else None
    except ValueError: return None

def calculate_age(value: Any, current_year: int | None = None) -> int | None:
    text = "" if value is None else str(value).strip()
    if len(text) < 4 or not text[:4].isdigit(): return None
    year = int(text[:4]); now = current_year or datetime.now().year
    age = now - year
    return age if 0 <= age <= 150 else None

def unique_output_path(folder: str | Path, stem: str) -> Path:
    folder = Path(folder); folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = folder / f"{stem}_{stamp}.xlsx"; counter = 1
    while path.exists():
        path = folder / f"{stem}_{stamp}_{counter}.xlsx"; counter += 1
    return path
