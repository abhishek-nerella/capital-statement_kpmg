"""
Arithmetic validation agent for the Capital Analysis Statement Generator.
Performs local Python cross-checks on each investor row, then sends all
numeric fields to Gemini for an independent second opinion.
Model: models/gemini-2.5-pro  |  max_output_tokens: 65536 (non-negotiable).
"""

from __future__ import annotations

import json

import pandas as pd
from google import genai as _genai

from audit_trail import log_event

_VALIDATION_MODEL = "models/gemini-2.5-pro"
_MAX_TOKENS = 65536


# ── Helpers ───────────────────────────────────────────────────────────────────

def _f(row: pd.Series, key: str) -> float:
    try:
        val = row.get(key, 0)
        return float(val) if val is not None and str(val).strip() != "" else 0.0
    except (TypeError, ValueError):
        return 0.0


# ── Local arithmetic checks ───────────────────────────────────────────────────

def _local_checks(row: pd.Series) -> list[dict]:
    """Compute four arithmetic checks in Python."""
    tev             = _f(row, "TEV")
    itd_dist        = _f(row, "INCEPTION_TO_DATE_DISTRIBUTION")
    closing_nav     = _f(row, "CLOSING_YTD_NAV")
    opening_nav     = _f(row, "OPENING_YTD_NAV")
    ytd_contrib     = _f(row, "YTD_CONTRIBUTION")
    ytd_dist        = _f(row, "YTD_DISTRIBUTION")
    inv_income      = _f(row, "INVESTMENT_INCOME")
    inv_expense     = _f(row, "INVESTMENT_EXPENSE")
    unrealized      = _f(row, "UNREALIZED_GAINS_LOSS")
    realized        = _f(row, "REALIZED_GAINS_LOSS")
    mgmt_fee        = _f(row, "MANAGEMENT_FEE")
    incentive_fee   = _f(row, "INCENTIVE_FEE")
    itd_contrib     = _f(row, "INCEPTION_TO_DATE_CONTRIBUTION")
    committed       = _f(row, "COMMITTED_CAPITAL")
    tev_ratio       = _f(row, "TEV_RATIO")

    checks: list[dict] = []

    # Check 1: TEV = ITD_DISTRIBUTION + CLOSING_NAV (tolerance ±$0.01)
    expected_tev = itd_dist + closing_nav
    delta1 = abs(tev - expected_tev)
    checks.append({
        "check": "TEV = INCEPTION_TO_DATE_DISTRIBUTION + CLOSING_YTD_NAV",
        "expected": expected_tev,
        "actual": tev,
        "delta": delta1,
        "pass": delta1 <= 0.01,
    })

    # Check 2: CLOSING_YTD_NAV ≈ NAV roll-forward (tolerance ±$1.00)
    expected_closing = (
        opening_nav + ytd_contrib - ytd_dist
        + (inv_income - inv_expense)
        + unrealized + realized
        - mgmt_fee - incentive_fee
    )
    delta2 = abs(closing_nav - expected_closing)
    checks.append({
        "check": "CLOSING_YTD_NAV = NAV roll-forward",
        "expected": expected_closing,
        "actual": closing_nav,
        "delta": delta2,
        "pass": delta2 <= 1.00,
    })

    # Check 3: TEV_RATIO ≈ TEV / ITD_CONTRIBUTION (tolerance ±0.01)
    expected_ratio = (tev / itd_contrib) if itd_contrib != 0.0 else 0.0
    delta3 = abs(tev_ratio - expected_ratio)
    checks.append({
        "check": "TEV_RATIO = TEV / INCEPTION_TO_DATE_CONTRIBUTION",
        "expected": expected_ratio,
        "actual": tev_ratio,
        "delta": delta3,
        "pass": delta3 <= 0.01,
    })

    # Check 4: REMAINING_COMMITMENT = COMMITTED - ITD_CONTRIB (flag if negative)
    remaining = committed - itd_contrib
    over_contrib = max(0.0, -remaining)  # positive only when negative remaining
    checks.append({
        "check": "REMAINING_COMMITMENT = COMMITTED_CAPITAL - INCEPTION_TO_DATE_CONTRIBUTION >= 0",
        "expected": committed,
        "actual": itd_contrib,
        "delta": over_contrib,
        "pass": remaining >= 0.0,
    })

    return checks


def _local_verdict(checks: list[dict]) -> str:
    """Derive verdict from local check results using the 5% relative-error rule."""
    failing = [c for c in checks if not c["pass"]]
    if not failing:
        return "ALL_PASS"
    for c in failing:
        delta = c.get("delta", 0)
        base  = abs(c.get("expected", 0))
        if not isinstance(delta, (int, float)):
            return "INVALID"
        pct = (delta / base) if base != 0 else float("inf")
        if pct > 0.05:
            return "INVALID"
    return "REVALUABLE"


# ── Gemini prompt ─────────────────────────────────────────────────────────────

def _build_prompt(row: pd.Series) -> str:
    fields = {
        "INVESTOR_NAME":                    str(row.get("INVESTOR_NAME", "UNKNOWN")),
        "COMMITTED_CAPITAL":                _f(row, "COMMITTED_CAPITAL"),
        "INCEPTION_TO_DATE_CONTRIBUTION":   _f(row, "INCEPTION_TO_DATE_CONTRIBUTION"),
        "INCEPTION_TO_DATE_DISTRIBUTION":   _f(row, "INCEPTION_TO_DATE_DISTRIBUTION"),
        "OPENING_YTD_NAV":                  _f(row, "OPENING_YTD_NAV"),
        "YTD_CONTRIBUTION":                 _f(row, "YTD_CONTRIBUTION"),
        "YTD_DISTRIBUTION":                 _f(row, "YTD_DISTRIBUTION"),
        "INVESTMENT_INCOME":                _f(row, "INVESTMENT_INCOME"),
        "INVESTMENT_EXPENSE":               _f(row, "INVESTMENT_EXPENSE"),
        "UNREALIZED_GAINS_LOSS":            _f(row, "UNREALIZED_GAINS_LOSS"),
        "REALIZED_GAINS_LOSS":              _f(row, "REALIZED_GAINS_LOSS"),
        "MANAGEMENT_FEE":                   _f(row, "MANAGEMENT_FEE"),
        "INCENTIVE_FEE":                    _f(row, "INCENTIVE_FEE"),
        "CLOSING_YTD_NAV":                  _f(row, "CLOSING_YTD_NAV"),
        "TEV":                              _f(row, "TEV"),
        "TEV_RATIO":                        _f(row, "TEV_RATIO"),
    }
    field_lines = "\n".join(f"  {k}: {v}" for k, v in fields.items())

    return f"""You are an arithmetic validation engine for private equity capital statements.

INVESTOR DATA:
{field_lines}

VALIDATION RULES:
1. TEV must equal INCEPTION_TO_DATE_DISTRIBUTION + CLOSING_YTD_NAV (tolerance: ±$0.01)
2. CLOSING_YTD_NAV must approximately equal OPENING_YTD_NAV + YTD_CONTRIBUTION - YTD_DISTRIBUTION + (INVESTMENT_INCOME - INVESTMENT_EXPENSE) + UNREALIZED_GAINS_LOSS + REALIZED_GAINS_LOSS - MANAGEMENT_FEE - INCENTIVE_FEE (tolerance: ±$1.00)
3. TEV_RATIO must approximately equal TEV / INCEPTION_TO_DATE_CONTRIBUTION (tolerance: ±0.01)
4. REMAINING_COMMITMENT implied = COMMITTED_CAPITAL - INCEPTION_TO_DATE_CONTRIBUTION — flag if negative

INSTRUCTIONS:
For each of the four checks above, re-derive the expected value from the raw numbers provided and state:
- PASS or FAIL
- expected value (numeric)
- actual value (numeric)
- delta (absolute difference, numeric)

Then state an overall verdict:
- "ALL_PASS": every check passes
- "REVALUABLE": at least one check fails but every failing delta is within 5% of its base value — flag and allow continuation; in "corrections" state which field to override and its corrected value
- "INVALID": any delta exceeds 5% of base value, or logically impossible values are present

Output ONLY valid JSON — no markdown fences, no preamble, no trailing text.

Required schema:
{{"verdict": "ALL_PASS" | "REVALUABLE" | "INVALID", "checks": [{{"check": "string", "pass": true|false, "expected": number, "actual": number, "delta": number}}], "corrections": {{}}, "notes": "string"}}"""


# ── Public API ────────────────────────────────────────────────────────────────

def validate_row(row: pd.Series, gemini_client, run_id: str) -> dict:
    """
    Validate the arithmetic fields of a single investor row.

    Performs local Python checks and calls Gemini (models/gemini-2.5-pro)
    for an independent second opinion.  Logs the result via audit_trail.
    Never raises — Gemini failures fall back to local-checks-only.

    Returns a dict with keys:
        investor, verdict, local_checks, gemini_checks,
        corrections, gemini_parse_error, notes
    """
    investor = str(row.get("INVESTOR_NAME", "UNKNOWN"))

    # ── Local checks ──────────────────────────────────────────────────────────
    local = _local_checks(row)
    verdict = _local_verdict(local)

    gemini_checks: list[dict] = []
    corrections:   dict       = {}
    gemini_parse_error        = False
    notes                     = ""

    # ── Gemini second opinion ─────────────────────────────────────────────────
    try:
        prompt = _build_prompt(row)
        response = gemini_client.models.generate_content(
            model=_VALIDATION_MODEL,
            config=_genai.types.GenerateContentConfig(
                max_output_tokens=_MAX_TOKENS,
            ),
            contents=prompt,
        )
        raw = response.text.strip()

        # Strip accidental markdown fences
        if raw.startswith("```"):
            lines = raw.splitlines()
            raw = "\n".join(lines[1:])
            if raw.endswith("```"):
                raw = raw[:-3].rstrip()

        gemini_result   = json.loads(raw)
        gemini_checks   = gemini_result.get("checks", [])
        corrections     = gemini_result.get("corrections", {})
        notes           = gemini_result.get("notes", "")
        gemini_verdict  = gemini_result.get("verdict", verdict)

        # Use the stricter of the two verdicts
        _severity = {"ALL_PASS": 0, "REVALUABLE": 1, "INVALID": 2}
        if _severity.get(gemini_verdict, 0) > _severity.get(verdict, 0):
            verdict = gemini_verdict

    except json.JSONDecodeError:
        gemini_parse_error = True
        notes = "Gemini response could not be parsed as JSON; falling back to local checks only."
    except Exception as exc:
        gemini_parse_error = True
        notes = f"Gemini call failed: {exc}; falling back to local checks only."

    result: dict = {
        "investor":          investor,
        "verdict":           verdict,
        "local_checks":      local,
        "gemini_checks":     gemini_checks,
        "corrections":       corrections,
        "gemini_parse_error": gemini_parse_error,
        "notes":             notes,
    }

    # ── Audit log ─────────────────────────────────────────────────────────────
    _event_map = {
        "ALL_PASS":    "validation_pass",
        "REVALUABLE":  "validation_revalued",
        "INVALID":     "validation_fail",
    }
    log_event(run_id, _event_map.get(verdict, "validation_fail"), investor, result)

    return result
