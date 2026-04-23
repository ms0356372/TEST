"""Application configuration constants."""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "ultrasound_checkin.db"

REQUIRED_FIELD_ALIASES = {
    "serial_no": ["序號"],
    "employee_id": ["人員工號", "工號"],
    "name": ["人員姓名", "姓名"],
    "gender": ["性別"],
    "timeslot": ["排程時段"],
    "item": ["項目"],
    "national_id": ["身分證", "身份證", "ID"],
}

STATUS_OPTIONS = [
    "上廁所",
    "心電圖",
    "先做其他",
    "超音波等候中",
    "超音波檢查",
]

GROUP_MAPPING = {
    "07:30~08:00": "A",
    "08:00~08:30": "B",
    "08:30~09:00": "C",
    "09:00~09:30": "D",
    "09:30~10:00": "E",
    "10:00~10:30": "F",
    "10:30~11:00": "G",
}
