"""Check-in and status services."""
from __future__ import annotations

import sqlite3
from typing import Optional

from app.config import STATUS_OPTIONS
from app.db import Database
from app.utils.time_slot import to_group_code


class CheckinError(Exception):
    """Check-in domain exception."""


class CheckinService:
    def __init__(self, db: Database) -> None:
        self.db = db

    def find_examinee(self, national_id: str) -> Optional[sqlite3.Row]:
        with self.db.connect() as conn:
            return conn.execute(
                """
                SELECT e.*, c.checkin_no, c.current_status
                FROM examinees e
                LEFT JOIN checkin_records c ON c.examinee_id = e.id
                WHERE e.national_id = ?
                """,
                (national_id.strip(),),
            ).fetchone()

    def check_in(self, national_id: str) -> str:
        with self.db.connect() as conn:
            examinee = conn.execute(
                "SELECT * FROM examinees WHERE national_id = ?",
                (national_id.strip(),),
            ).fetchone()
            if not examinee:
                raise CheckinError("查無此身分證資料，請先匯入名單。")

            existing = conn.execute(
                "SELECT checkin_no FROM checkin_records WHERE examinee_id = ?",
                (examinee["id"],),
            ).fetchone()
            if existing:
                raise CheckinError(f"此人已報到，報到編號：{existing['checkin_no']}")

            group_code = to_group_code(examinee["timeslot"])
            sequence_row = conn.execute(
                "SELECT last_sequence FROM group_sequences WHERE group_code = ?",
                (group_code,),
            ).fetchone()
            if not sequence_row:
                raise CheckinError(f"找不到組別 {group_code} 的流水號設定")

            next_no = int(sequence_row["last_sequence"]) + 1
            checkin_no = f"{group_code}{next_no}"
            now = self.db.now_iso()

            conn.execute(
                """
                UPDATE group_sequences
                SET last_sequence = ?, updated_at = ?
                WHERE group_code = ?
                """,
                (next_no, now, group_code),
            )
            conn.execute(
                """
                INSERT INTO checkin_records(
                    examinee_id, group_code, sequence_no, checkin_no,
                    checked_in_at, current_status
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (examinee["id"], group_code, next_no, checkin_no, now, ""),
            )
            return checkin_no

    def update_status(self, checkin_record_id: int, status: str) -> None:
        if status not in STATUS_OPTIONS:
            raise CheckinError("無效狀態。")

        now = self.db.now_iso()
        with self.db.connect() as conn:
            found = conn.execute(
                "SELECT id FROM checkin_records WHERE id = ?",
                (checkin_record_id,),
            ).fetchone()
            if not found:
                raise CheckinError("尚未報到，無法更新狀態。")

            conn.execute(
                "UPDATE checkin_records SET current_status = ? WHERE id = ?",
                (status, checkin_record_id),
            )
            conn.execute(
                """
                INSERT INTO status_history(checkin_record_id, status, changed_at)
                VALUES (?, ?, ?)
                """,
                (checkin_record_id, status, now),
            )
