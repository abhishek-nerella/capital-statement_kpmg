"""
Audit trail module for the Capital Analysis Statement Generator.
Writes a structured, newline-delimited JSON log (JSONL) for every
generation run, and in parallel a plain-English, human-readable log of
the same events. All public functions are thread-safe and never raise
exceptions — errors are printed to stderr only.
"""

from __future__ import annotations

import json
import sys
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path

_LOG_PATH       = Path("/tmp/audit_log.jsonl") if Path("/tmp").exists() else Path("audit_log.jsonl")
_PLAIN_LOG_PATH = Path("/tmp/audit_log.txt")   if Path("/tmp").exists() else Path("audit_log.txt")
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


def _append_plain(line: str) -> None:
    try:
        with _lock:
            with _PLAIN_LOG_PATH.open("a", encoding="utf-8") as fh:
                fh.write(line + "\n")
    except Exception as exc:
        print(f"[audit_trail] plain-log write error: {exc}", file=sys.stderr)


def _local_ts(iso_utc: str) -> str:
    """Trim an ISO-8601 UTC timestamp down to 'YYYY-MM-DD HH:MM:SS UTC' for readability."""
    return iso_utc[:19].replace("T", " ") + " UTC"


def _plain_event_sentence(event_type: str, investor: str, detail: dict) -> str:
    """Translate one structured audit event into a single plain-English sentence."""
    detail = detail or {}

    if event_type in ("validation_pass", "hf_validation_pass"):
        notes = detail.get("notes")
        return f"{investor}: validation passed — all arithmetic checks OK." + (f" {notes}" if notes else "")

    if event_type in ("validation_revalued", "hf_validation_flagged"):
        corrections = detail.get("corrections") or {}
        notes = detail.get("notes", "")
        if corrections:
            fields = ", ".join(corrections.keys())
            return f"{investor}: validation flagged for revaluation — corrected field(s): {fields}. {notes}".strip()
        return f"{investor}: validation flagged for revaluation. {notes}".strip()

    if event_type in ("validation_fail", "hf_validation_fail"):
        notes = detail.get("notes", "")
        reason = detail.get("reason", "")
        return f"{investor}: validation FAILED — {reason or notes or 'see notes'}".strip()

    if event_type == "document_generated":
        fname = detail.get("file", "")
        verdict = detail.get("verdict")
        suffix = f" (verdict: {verdict})" if verdict else ""
        return f"{investor}: statement generated successfully — {fname}{suffix}"

    if event_type == "document_failed":
        reason = detail.get("reason", "unknown error")
        return f"{investor}: statement generation FAILED — {reason}"

    if event_type == "wrangler_change":
        col     = detail.get("column", "?")
        action  = detail.get("action", "changed")
        orig    = detail.get("original_value")
        coerced = detail.get("coerced_value")
        return f"{investor}: column '{col}' {action} — '{orig}' became '{coerced}'."

    if event_type == "gemini_insight":
        return f"{investor}: AI insight generated."

    # Fallback: render whatever key/value pairs are present
    parts = "; ".join(f"{k}: {v}" for k, v in detail.items())
    return f"{investor}: {event_type} — {parts}" if parts else f"{investor}: {event_type}"


def start_run(input_filename: str, investor_count: int) -> str:
    """
    Record the start of a generation run.

    Returns a UUID4 run_id that must be passed to log_event and close_run.
    On internal error, a fresh UUID is still returned so callers always
    receive a valid run_id.
    """
    try:
        run_id = str(uuid.uuid4())
        ts = _now_iso()
        _append({
            "run_id": run_id,
            "timestamp_utc": ts,
            "input_file": input_filename,
            "investor_count": investor_count,
            "status": "started",
        })
        _append_plain(
            f"[{_local_ts(ts)}] Run {run_id} started — "
            f"{investor_count} investor(s) queued from '{input_filename}'."
        )
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
        ts = _now_iso()
        _append({
            "run_id": run_id,
            "timestamp_utc": ts,
            "event_type": event_type,
            "investor": investor,
            "detail": detail,
        })
        _append_plain(f"[{_local_ts(ts)}] " + _plain_event_sentence(event_type, investor, detail))
    except Exception as exc:
        print(f"[audit_trail] log_event error: {exc}", file=sys.stderr)


def close_run(run_id: str, success_count: int, fail_count: int) -> None:
    """Record the completion of a generation run with final success/fail counts."""
    try:
        ts = _now_iso()
        _append({
            "run_id": run_id,
            "timestamp_utc": ts,
            "status": "completed",
            "success_count": success_count,
            "fail_count": fail_count,
        })
        _append_plain(
            f"[{_local_ts(ts)}] Run {run_id} completed — "
            f"{success_count} succeeded, {fail_count} failed."
        )
    except Exception as exc:
        print(f"[audit_trail] close_run error: {exc}", file=sys.stderr)
