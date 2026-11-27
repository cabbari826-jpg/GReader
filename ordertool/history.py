from __future__ import annotations
import sqlite3
from pathlib import Path
from datetime import datetime


class HistoryStore:
    def __init__(self, database_path: str | Path):
        self.path = Path(database_path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _connect(self):
        conn = sqlite3.connect(self.path, timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        with self._connect() as conn:
            conn.execute("""CREATE TABLE IF NOT EXISTS processed_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_hash TEXT NOT NULL UNIQUE,
                filename TEXT NOT NULL,
                source_path TEXT NOT NULL,
                order_number TEXT,
                processed_at TEXT NOT NULL,
                conversion_status TEXT NOT NULL,
                validation_status TEXT NOT NULL,
                discrepancy_count INTEGER NOT NULL DEFAULT 0,
                output_file TEXT,
                validation_report TEXT,
                photo_library TEXT
            )""")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_processed_at ON processed_files(processed_at DESC)")

    def contains_hash(self, file_hash: str) -> bool:
        with self._connect() as conn:
            return conn.execute("SELECT 1 FROM processed_files WHERE file_hash=?", (file_hash,)).fetchone() is not None

    def record(self, order, validation_report="", photo_library=""):
        discrepancy_count = sum(p.validation_status == "DISCREPANCY" for p in order.products) + len(order.order_issues)
        with self._connect() as conn:
            conn.execute("""INSERT INTO processed_files
                (file_hash, filename, source_path, order_number, processed_at, conversion_status,
                 validation_status, discrepancy_count, output_file, validation_report, photo_library)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(file_hash) DO UPDATE SET
                    processed_at=excluded.processed_at, conversion_status=excluded.conversion_status,
                    validation_status=excluded.validation_status, discrepancy_count=excluded.discrepancy_count,
                    output_file=excluded.output_file, validation_report=excluded.validation_report,
                    photo_library=excluded.photo_library""",
                (order.source_hash, Path(order.source_pdf).name, order.source_pdf, order.order_number,
                 order.processed_at or datetime.now().isoformat(timespec="seconds"), order.conversion_status,
                 order.validation_status, discrepancy_count, order.order_workbook, validation_report, photo_library))

    def list_recent(self, limit=500):
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM processed_files ORDER BY processed_at DESC LIMIT ?", (limit,)).fetchall()
            return [dict(r) for r in rows]
