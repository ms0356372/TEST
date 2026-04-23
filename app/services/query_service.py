"""List query and filtering service."""
from __future__ import annotations

from typing import List, Optional

from app.db import Database


class QueryService:
    def __init__(self, db: Database) -> None:
        self.db = db

    def list_examinees(
        self,
        status_filter: str = "all",
        group_filter: str = "all",
        keyword: str = "",
    ) -> List[dict]:
        sql = """
            SELECT
                e.id,
                e.serial_no,
                e.employee_id,
                e.name,
                e.gender,
                e.timeslot,
                e.item,
                e.national_id,
                c.id AS checkin_record_id,
                c.group_code,
                c.sequence_no,
                c.checkin_no,
                c.current_status,
                c.checked_in_at
            FROM examinees e
            LEFT JOIN checkin_records c ON c.examinee_id = e.id
            WHERE 1=1
        """
        params: list[object] = []

        if status_filter == "checked_in":
            sql += " AND c.id IS NOT NULL"
        elif status_filter == "not_checked_in":
            sql += " AND c.id IS NULL"

        if group_filter != "all":
            sql += " AND c.group_code = ?"
            params.append(group_filter)

        if keyword.strip():
            sql += " AND (e.name LIKE ? OR e.employee_id LIKE ? OR e.national_id LIKE ?)"
            pattern = f"%{keyword.strip()}%"
            params.extend([pattern, pattern, pattern])

        sql += " ORDER BY e.serial_no ASC"

        with self.db.connect() as conn:
            rows = conn.execute(sql, params).fetchall()
            return [dict(row) for row in rows]

    def get_checkin_record_id_by_examinee(self, examinee_id: int) -> Optional[int]:
        with self.db.connect() as conn:
            row = conn.execute(
                "SELECT id FROM checkin_records WHERE examinee_id = ?",
                (examinee_id,),
            ).fetchone()
            return int(row["id"]) if row else None
