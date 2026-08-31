"""應用程式檔案記錄器。"""
import logging
from datetime import datetime
from config_manager import get_app_dir


def get_logger() -> logging.Logger:
    """建立每日 UTF-8 log，避免重複掛載 handler。"""
    logger = logging.getLogger("ExcelTransformTool")
    if not logger.handlers:
        folder = get_app_dir() / "logs"; folder.mkdir(parents=True, exist_ok=True)
        handler = logging.FileHandler(folder / f"{datetime.now():%Y%m%d}.log", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(handler); logger.setLevel(logging.INFO)
    return logger
