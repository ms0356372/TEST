"""Time slot and group conversion helpers."""
from app.config import GROUP_MAPPING


def to_group_code(timeslot: str) -> str:
    normalized = timeslot.strip()
    if normalized not in GROUP_MAPPING:
        raise ValueError(f"不支援的排程時段：{timeslot}")
    return GROUP_MAPPING[normalized]
