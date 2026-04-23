"""Database access layer for SQLite."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Generator

from app.config import DB_PATH


class Database:
    """Simple database helper wrapping sqlite3 connection handling."""

    def __init__(self, db_path: Path = DB_PATH) -> None:
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_schema()

    @contextmanager
    def connect(self) -> Generator[sqlite3.Connection, None, None]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _init_schema(self) -> None:
        with self.connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS examinees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    serial_no TEXT NOT NULL,
                    employee_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    gender TEXT NOT NULL,
                    timeslot TEXT NOT NULL,
                    item TEXT NOT NULL,
                    national_id TEXT NOT NULL UNIQUE,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS checkin_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    examinee_id INTEGER NOT NULL UNIQUE,
                    group_code TEXT NOT NULL,
                    sequence_no INTEGER NOT NULL,
                    checkin_no TEXT NOT NULL UNIQUE,
                    checked_in_at TEXT NOT NULL,
                    current_status TEXT,
                    FOREIGN KEY(examinee_id) REFERENCES examinees(id)
                );

                CREATE TABLE IF NOT EXISTS group_sequences (
                    group_code TEXT PRIMARY KEY,
                    last_sequence INTEGER NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS status_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    checkin_record_id INTEGER NOT NULL,
                    status TEXT NOT NULL,
                    changed_at TEXT NOT NULL,
                    FOREIGN KEY(checkin_record_id) REFERENCES checkin_records(id)
                );

                CREATE TABLE IF NOT EXISTS system_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                """
            )

            now = datetime.now().isoformat(timespec="seconds")
            for code in ["A", "B", "C", "D", "E", "F", "G"]:
                conn.execute(
                    """
                    INSERT OR IGNORE INTO group_sequences(group_code, last_sequence, updated_at)
                    VALUES (?, 0, ?)
                    """,
                    (code, now),
                )

    @staticmethod
    def now_iso() -> str:
        return datetime.now().isoformat(timespec="seconds")
