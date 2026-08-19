"""共用輸入、輸出路徑驗證。"""
from __future__ import annotations

import os
from pathlib import Path

from core.exceptions import AnalysisError


def validate_excel_input(path: str | Path, template_name: str) -> Path:
    """確認輸入是存在且可讀取的 xlsx 一般檔案。"""
    candidate = Path(path).expanduser()
    if not candidate.is_file():
        raise AnalysisError(f"{template_name}檔案不存在或不是一般檔案。")
    if candidate.suffix.lower() != ".xlsx":
        raise AnalysisError(f"{template_name}必須是 .xlsx 檔案。")
    if not os.access(candidate, os.R_OK):
        raise AnalysisError(f"{template_name}檔案無法讀取。")
    return candidate.resolve()


def validate_output_directory(path: str | Path) -> Path:
    """確認輸出資料夾存在且可用實際暫存檔測試寫入權限。"""
    candidate = Path(path).expanduser()
    if not candidate.is_dir():
        raise AnalysisError("輸出資料夾不存在或不是資料夾。")

    probe = candidate / ".excel_health_tool_write_test"
    try:
        with probe.open("w", encoding="utf-8") as stream:
            stream.write("ok")
        probe.unlink()
    except OSError as exc:
        try:
            probe.unlink(missing_ok=True)
        except OSError:
            pass
        raise AnalysisError(f"輸出資料夾無法寫入：{exc}") from exc
    return candidate.resolve()
