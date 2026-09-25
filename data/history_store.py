"""Tiny SQLite wrapper for comparison history."""

import os
import sqlite3

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.db")


def _connect():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            meter_number TEXT,
            old_file TEXT,
            new_file TEXT,
            old_firmware TEXT,
            new_firmware TEXT,
            total_changes INTEGER,
            high_changes INTEGER,
            medium_changes INTEGER,
            low_changes INTEGER,
            report_path TEXT,
            status TEXT
        )
    """)
    return conn


def add_entry(result, status="SUCCESS"):
    conn = _connect()
    conn.execute(
        """INSERT INTO history
           (timestamp, meter_number, old_file, new_file, old_firmware, new_firmware,
            total_changes, high_changes, medium_changes, low_changes, report_path, status)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            result.get("timestamp"), result.get("meter_number"),
            result.get("old_file"), result.get("new_file"),
            result.get("old_firmware"), result.get("new_firmware"),
            len(result.get("changes", [])), result.get("high", 0),
            result.get("medium", 0), result.get("low", 0),
            result.get("output_path"), status,
        ),
    )
    conn.commit()
    conn.close()


def add_failure(old_file, new_file, error_message, timestamp):
    conn = _connect()
    conn.execute(
        """INSERT INTO history
           (timestamp, meter_number, old_file, new_file, old_firmware, new_firmware,
            total_changes, high_changes, medium_changes, low_changes, report_path, status)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (timestamp, "", old_file or "", new_file or "", "", "", 0, 0, 0, 0, "",
         f"FAILED: {error_message}"),
    )
    conn.commit()
    conn.close()


def all_entries():
    conn = _connect()
    cur = conn.execute("SELECT * FROM history ORDER BY id DESC")
    columns = [d[0] for d in cur.description]
    rows = [dict(zip(columns, row)) for row in cur.fetchall()]
    conn.close()
    return rows


def stats():
    entries = all_entries()
    total = len(entries)
    success = sum(1 for e in entries if e["status"] == "SUCCESS")
    failed = total - success
    return {"total": total, "success": success, "failed": failed}
