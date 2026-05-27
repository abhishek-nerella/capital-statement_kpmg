"""
Audit trail module for the Capital Analysis Statement Generator.
Writes a structured, newline-delimited JSON log (JSONL) for every
generation run. All public functions are thread-safe and never raise
exceptions — errors are printed to stderr only.
"""

from __future__ import annotations

import json
import sys
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path

_LOG_PATH = Path("audit_log.jsonl")
_lock = threading.Lock()


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _append(record: dict) -> None:
    try:
        with _lock:
            with _LOG_PATH.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps(record, default=str) + "\n")
    except Exception as exc:
        print(f"[audit_trail] write error: {exc}", file=sys.stderr)


def start_run(input_filename: str, investor_count: int) -> str:
    """
    Record the start of a generation run.

    Returns a UUID4 run_id that must be passed to log_event and close_run.
    On internal error, a fresh UUID is still returned so callers always
    receive a valid run_id.
    """
    try:
        run_id = str(uuid.uuid4())
        _append({
            "run_id": run_id,
            "timestamp_utc": _now_iso(),
            "input_file": input_filename,
            "investor_count": investor_count,
            "status": "started",
        })
        return run_id
    except Exception as exc:
        print(f"[audit_trail] start_run error: {exc}", file=sys.stderr)
        return str(uuid.uuid4())


def log_event(run_id: str, event_type: str, investor: str, detail: dict) -> None:
    """
    Append a single structured event line to the audit log.

    Recognised event_type values:
        wrangler_change, validation_pass, validation_fail,
        validation_revalued, document_generated, document_failed,
        gemini_insight
    """
    try:
        _append({
            "run_id": run_id,
            "timestamp_utc": _now_iso(),
            "event_type": event_type,
            "investor": investor,
            "detail": detail,
        })
    except Exception as exc:
        print(f"[audit_trail] log_event error: {exc}", file=sys.stderr)


def close_run(run_id: str, success_count: int, fail_count: int) -> None:
    """Record the completion of a generation run with final success/fail counts."""
    try:
        _append({
            "run_id": run_id,
            "timestamp_utc": _now_iso(),
            "status": "completed",
            "success_count": success_count,
            "fail_count": fail_count,
        })
    except Exception as exc:
        print(f"[audit_trail] close_run error: {exc}", file=sys.stderr)
