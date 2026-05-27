"""
Data preprocessing module for the Capital Analysis Statement Generator.
Normalises raw investor data before document generation and returns a
structured log of every transformation applied.
"""

from __future__ import annotations

import pandas as pd

from generate_capital_statements import REQUIRED_COLS

# Numeric columns within REQUIRED_COLS (excludes dates, text identifiers)
_NUMERIC_REQUIRED: set[str] = {
    "COMMITTED_CAPITAL",
    "INCEPTION_TO_DATE_CONTRIBUTION",
    "INCEPTION_TO_DATE_DISTRIBUTION",
    "OPENING_YTD_NAV",
    "YTD_CONTRIBUTION",
    "YTD_DISTRIBUTION",
    "INVESTMENT_INCOME",
    "INVESTMENT_EXPENSE",
    "UNREALIZED_GAINS_LOSS",
    "REALIZED_GAINS_LOSS",
    "MANAGEMENT_FEE",
    "INCENTIVE_FEE",
    "CLOSING_YTD_NAV",
    "TEV",
    "TEV_RATIO",
}


def wrangle(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """
    Clean and normalise a raw investor DataFrame.

    Transformations applied in order:
    1. Column name normalisation (strip + uppercase)
    2. String whitespace cleanup
    3. Date normalisation for FROM_DATE / TO_DATE
    4. Currency-symbol stripping and float cast for all numeric REQUIRED_COLS
    5. NaN fill for MANAGEMENT_FEE and INCENTIVE_FEE → 0.0
    6. NaN fill for all remaining numeric REQUIRED_COLS → 0.0

    Returns a cleaned *copy* of the input DataFrame and a list of
    WranglerEvent dicts (keys: column, investor, original_value,
    coerced_value, action).
    """
    df = df.copy()
    events: list[dict] = []

    def _investor(idx) -> str:
        if "INVESTOR_NAME" in df.columns:
            val = df.at[idx, "INVESTOR_NAME"]
            return str(val).strip() if pd.notna(val) else "UNKNOWN"
        return "UNKNOWN"

    def _event(col: str, idx, original, coerced, action: str) -> None:
        events.append({
            "column": col,
            "investor": _investor(idx),
            "original_value": original,
            "coerced_value": coerced,
            "action": action,
        })

    # ── Step 1: Column name normalisation ────────────────────────────────────
    df.columns = [str(c).strip().upper() for c in df.columns]

    # ── Step 2: String whitespace cleanup ────────────────────────────────────
    for col in df.select_dtypes(include=["object"]).columns:
        for idx in df.index:
            original = df.at[idx, col]
            if isinstance(original, str):
                cleaned = original.strip()
                if cleaned != original:
                    _event(col, idx, original, cleaned, "stripped leading/trailing whitespace")
                    df.at[idx, col] = cleaned

    # ── Step 3: Date normalisation ────────────────────────────────────────────
    for date_col in ("FROM_DATE", "TO_DATE"):
        if date_col not in df.columns:
            continue
        original_series = df[date_col].copy()
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        for idx in df.index:
            orig_val = original_series.at[idx]
            if pd.isnull(df.at[idx, date_col]) and pd.notna(orig_val):
                _event(
                    date_col, idx, str(orig_val), None,
                    f"date parse failed for {date_col}; coerced to NaT",
                )

    # ── Step 4: Currency parsing for numeric REQUIRED_COLS ────────────────────
    # Phase 1: generate per-cell events for string values only.
    # Phase 2: replace the entire column with pd.to_numeric so pandas handles
    # the dtype transition (avoids ArrowStringArray → float assignment errors).
    numeric_present = _NUMERIC_REQUIRED & set(df.columns)
    for col in numeric_present:
        for idx in df.index:
            original = df.at[idx, col]
            if isinstance(original, str):
                cleaned = original.replace("$", "").replace(",", "").replace("%", "").strip()
                try:
                    value = float(cleaned)
                    _event(col, idx, original, value,
                           "stripped currency symbols and cast to float")
                except (ValueError, TypeError):
                    _event(col, idx, original, 0.0,
                           f"could not cast '{original}' to float; defaulted to 0.0")
        # Vectorised column replacement — safe for any backing dtype
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(r"[$,%]", "", regex=True).str.strip(),
            errors="coerce",
        ).fillna(0.0)

    # ── Step 5: Null fee defaults ─────────────────────────────────────────────
    for fee_col in ("MANAGEMENT_FEE", "INCENTIVE_FEE"):
        if fee_col not in df.columns:
            continue
        for idx in df.index:
            if pd.isnull(df.at[idx, fee_col]):
                _event(fee_col, idx, None, 0.0, f"filled NaN {fee_col} with 0.0")
        df[fee_col] = df[fee_col].fillna(0.0)

    # ── Step 6: Null numeric defaults for remaining REQUIRED_COLS ─────────────
    for col in numeric_present - {"MANAGEMENT_FEE", "INCENTIVE_FEE"}:
        for idx in df.index:
            if pd.isnull(df.at[idx, col]):
                _event(col, idx, None, 0.0, f"filled NaN {col} with 0.0")
        df[col] = df[col].fillna(0.0)

    return df, events
