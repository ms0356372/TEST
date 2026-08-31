"""Windows EXE 自動建置器。

BAT 僅負責尋找 Python 與保留視窗；實際流程放在 Python，避免 cmd 對中文路徑、
括號及 ERRORLEVEL 的解析差異造成雙擊後無訊息結束。
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

from version import APP_NAME, APP_VERSION

ROOT = Path(__file__).resolve().parent
LOG_PATH = ROOT / "build_log.txt"
EXE_PATH = ROOT / "dist" / f"{APP_NAME}.exe"
RELEASE_PATH = ROOT / "release" / f"{APP_NAME}_v{APP_VERSION}.exe"
SPEC_PATH = ROOT / f"{APP_NAME}.spec"


class BuildFailure(RuntimeError):
    """代表某個建置階段未通過。"""


def show(message: str = "") -> None:
    """同時顯示並立即寫入建置紀錄。"""
    print(message, flush=True)
    with LOG_PATH.open("a", encoding="utf-8") as log:
        log.write(message + "\n")


def run(command: list[str], stage: str) -> None:
    """即時轉送子程序輸出，失敗時中止後續建置。"""
    show(f"COMMAND: {subprocess.list2cmdline(command)}")
    with LOG_PATH.open("a", encoding="utf-8") as log:
        process = subprocess.Popen(
            command,
            cwd=ROOT,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        assert process.stdout is not None
        for line in process.stdout:
            print(line, end="", flush=True)
            log.write(line)
        code = process.wait()
    if code:
        raise BuildFailure(f"{stage} 失敗（exit code {code}）")


def remove_old_outputs() -> None:
    """只清除本專案已知的 PyInstaller 與 release 產物。"""
    for folder in (ROOT / "build", ROOT / "dist"):
        if folder.exists():
            shutil.rmtree(folder)
    if SPEC_PATH.exists():
        SPEC_PATH.unlink()
    # 舊版 release 不可在本次 EXE 驗證前繼續冒充新成品。
    if RELEASE_PATH.exists():
        RELEASE_PATH.unlink()
    release_dir = RELEASE_PATH.parent
    if release_dir.exists() and not any(release_dir.iterdir()):
        release_dir.rmdir()


def verify_executable() -> None:
    """啟動 windowed EXE，等待其將完整 Self Test 寫入暫存文字檔。"""
    result_path = Path(tempfile.gettempdir()) / "excel_tool_self_test.txt"
    result_path.unlink(missing_ok=True)
    subprocess.Popen([str(EXE_PATH), "--self-test", "--result-file", str(result_path)], cwd=ROOT)
    deadline = time.monotonic() + 90
    while time.monotonic() < deadline and not result_path.exists():
        time.sleep(0.25)
    if not result_path.exists():
        raise BuildFailure("EXE Self Test 逾時，未產生結果檔")
    content = result_path.read_text(encoding="utf-8-sig")
    show(content.rstrip())
    if not content.splitlines() or content.splitlines()[0].strip() != "PASS":
        raise BuildFailure("EXE Self Test 回報 FAIL")


def build(check_only: bool = False) -> None:
    """依指定順序測試、打包、驗證，最後才發布 EXE。"""
    LOG_PATH.write_text(
        f"Build started: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"Working directory: {ROOT}\n",
        encoding="utf-8",
    )
    python = sys.executable
    show("[1/9] 檢查 Python...")
    show(f"[成功] {sys.version.replace(os.linesep, ' ')}")
    show("[2/9] 檢查 pip...")
    run([python, "-m", "pip", "--version"], "pip 檢查")
    show("[3/9] 安裝/確認必要套件...")
    run([python, "-m", "pip", "install", "-r", "requirements.txt"], "套件安裝")
    run([python, "-c", "import openpyxl, PyInstaller; print('openpyxl / PyInstaller: PASS')"], "套件匯入")
    show("[4/9] 執行核心單元測試...")
    run([python, "test_transform.py"], "核心單元測試")
    show("[5/9] 執行 Python Self Test...")
    run([python, "main.py", "--self-test"], "Python Self Test")
    if check_only:
        show("CHECK ONLY: PASS")
        return
    if os.name != "nt":
        raise BuildFailure("PyInstaller Windows EXE 必須在 Windows 執行")
    show("[6/9] 清除舊版 Build...")
    remove_old_outputs()
    show("[成功] 舊版輸出已清除")
    show("[7/9] 建立 EXE...")
    run([python, "-m", "PyInstaller", "--onefile", "--windowed", "--clean", "--noconfirm", "--name", APP_NAME, "main.py"], "PyInstaller")
    show("[8/9] 檢查 EXE...")
    if not EXE_PATH.is_file() or EXE_PATH.stat().st_size == 0:
        raise BuildFailure(f"找不到有效 EXE：{EXE_PATH}")
    show(f"[成功] {EXE_PATH}")
    show("[9/9] 執行 EXE Self Test...")
    verify_executable()
    RELEASE_PATH.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(EXE_PATH, RELEASE_PATH)
    size_mb = RELEASE_PATH.stat().st_size / 1024 / 1024
    show("=" * 40)
    show("BUILD SUCCESS")
    show(f"版本：v{APP_VERSION}")
    show(f"正式 EXE：{RELEASE_PATH}")
    show(f"EXE 大小：{size_mb:.1f} MB")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check-only", action="store_true", help="只跑打包前檢查")
    args = parser.parse_args()
    try:
        build(args.check_only)
        return 0
    except Exception as exc:
        show(f"[失敗] {exc}")
        show(f"詳細資訊：{LOG_PATH}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
