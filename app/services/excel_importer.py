"""Excel import service with column normalization."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List

import pandas as pd

from app.config import REQUIRED_FIELD_ALIASES
from app.db import Database


class ImportValidationError(Exception):
    """Raised when Excel headers or data are invalid."""


@dataclass
class ImportResult:
    imported_count: int
    updated_count: int


class ExcelImporter:
    def __init__(self, db: Database) -> None:
        self.db = db

    def import_file(self, path: str) -> ImportResult:
        df = pd.read_excel(path, dtype=str)
        df.columns = [str(col).strip() for col in df.columns]

        col_map = self._normalize_columns(df.columns.tolist())
        renamed = df.rename(columns={src: dst for dst, src in col_map.items()})
        required = list(REQUIRED_FIELD_ALIASES.keys())
        working = renamed[required].fillna("").astype(str)

        for col in required:
            if (working[col].str.strip() == "").any():
                raise ImportValidationError(f"欄位「{col}」有空值，請修正後再匯入。")

        imported_count = 0
        updated_count = 0
        now = self.db.now_iso()

        with self.db.connect() as conn:
            for _, row in working.iterrows():
                payload = {k: row[k].strip() for k in required}
                existing = conn.execute(
                    "SELECT id FROM examinees WHERE national_id = ?",
                    (payload["national_id"],),
                ).fetchone()

                if existing:
                    conn.execute(
                        """
                        UPDATE examinees
                        SET serial_no = ?, employee_id = ?, name = ?, gender = ?,
                            timeslot = ?, item = ?, updated_at = ?
                        WHERE id = ?
                        """,
                        (
                            payload["serial_no"],
                            payload["employee_id"],
                            payload["name"],
                            payload["gender"],
                            payload["timeslot"],
                            payload["item"],
                            now,
                            existing["id"],
                        ),
                    )
                    updated_count += 1
                else:
                    conn.execute(
                        """
                        INSERT INTO examinees(
                            serial_no, employee_id, name, gender, timeslot,
                            item, national_id, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            payload["serial_no"],
                            payload["employee_id"],
                            payload["name"],
                            payload["gender"],
                            payload["timeslot"],
                            payload["item"],
                            payload["national_id"],
                            now,
                            now,
                        ),
                    )
                    imported_count += 1

        return ImportResult(imported_count=imported_count, updated_count=updated_count)

    def _normalize_columns(self, headers: List[str]) -> Dict[str, str]:
        normalized = {}
        missing = []

        for target, aliases in REQUIRED_FIELD_ALIASES.items():
            matched = next((h for h in headers if h in aliases), None)
            if not matched:
                missing.append(f"{target}（可用欄位：{' / '.join(aliases)}）")
            else:
                normalized[target] = matched

        if missing:
            raise ImportValidationError("缺少必要欄位：\n" + "\n".join(missing))

        return normalized
