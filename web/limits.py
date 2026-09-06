"""Per-visitor quotas and the one-export-at-a-time gate.

Two different problems, deliberately solved two different ways:

* **Quota** is about fairness over hours, so it has to survive a worker
  recycle (`--max-requests` restarts this process regularly) and any number of
  workers. It lives in SQLite.
* **Concurrency** is about this 2GB box surviving a single export, which peaks
  around 350MB. That is a within-process question, so it is a semaphore.
"""

from __future__ import annotations

import hashlib
import sqlite3
import threading
import time
from pathlib import Path

# One export at a time. The page itself stays responsive because gunicorn runs
# threaded workers -- it is only the memory-hungry part that queues.
_export_gate = threading.BoundedSemaphore(1)

EXPORT_WAIT_SECONDS = 8


class Busy(Exception):
    """Another export is already running and did not finish in time."""


class ExportSlot:
    """Context manager around the single export slot."""

    def __enter__(self) -> "ExportSlot":
        if not _export_gate.acquire(timeout=EXPORT_WAIT_SECONDS):
            raise Busy()
        return self

    def __exit__(self, *_exc_info: object) -> None:
        _export_gate.release()


class Quota:
    """Sliding-window request counter keyed by a hashed visitor identity."""

    def __init__(self, db_path: Path, windows: dict[str, tuple[int, int]]) -> None:
        """`windows` maps a name to (max_events, window_seconds)."""
        self.db_path = db_path
        self.windows = windows
        self._lock = threading.Lock()
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            connection.execute(
                "CREATE TABLE IF NOT EXISTS events ("
                "  subject TEXT NOT NULL,"
                "  kind TEXT NOT NULL,"
                "  at REAL NOT NULL"
                ")"
            )
            connection.execute("CREATE INDEX IF NOT EXISTS events_lookup ON events (subject, kind, at)")

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path, timeout=5)
        connection.execute("PRAGMA journal_mode=WAL")
        return connection

    @staticmethod
    def subject(client_ip: str | None) -> str:
        """Hash the address so the quota table never holds a visitor's IP."""
        return hashlib.sha256(f"topdf:{client_ip or 'unknown'}".encode()).hexdigest()[:32]

    def check_and_record(self, subject: str, kind: str) -> tuple[bool, str | None]:
        """Record one event if every window still has room.

        Returns (allowed, message). Fails **open** on a database error: a
        broken counter should not take the tool down, and the concurrency gate
        plus nginx's own rate limit are still in front of the expensive work.
        """
        now = time.time()
        longest = max(seconds for _limit, seconds in self.windows.values())
        try:
            with self._lock, self._connect() as connection:
                connection.execute(
                    "DELETE FROM events WHERE at < ?", (now - longest,)
                )
                for name, (limit, seconds) in self.windows.items():
                    (used,) = connection.execute(
                        "SELECT COUNT(*) FROM events WHERE subject = ? AND kind = ? AND at >= ?",
                        (subject, kind, now - seconds),
                    ).fetchone()
                    if used >= limit:
                        return False, _limit_message(limit, seconds)
                connection.execute(
                    "INSERT INTO events (subject, kind, at) VALUES (?, ?, ?)",
                    (subject, kind, now),
                )
            return True, None
        except sqlite3.Error:
            return True, None


def _limit_message(limit: int, seconds: int) -> str:
    if seconds >= 86400:
        period = "a day"
    elif seconds >= 3600:
        hours = seconds // 3600
        period = "an hour" if hours == 1 else f"{hours} hours"
    else:
        period = f"{seconds // 60} minutes"
    plural = "" if limit == 1 else "s"
    return f"You've reached the limit of {limit} catalog{plural} {period}. Try again later."
