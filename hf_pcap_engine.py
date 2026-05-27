"""
Hedge Fund PCAP Reader — parses the KPMG pre-calculated PCAP Excel format.

The Excel has one data row per investor. Row 1 is merged-cell group headers
("LOCK-UP & REDEMPTION TERMS", "SIDE LETTER TERMS", "WATERFALL & FEE PARAMETERS").
Row 2 is the actual column names. Data starts from row 3.

Column groups:
  Unit Prices | CQ $ | CQ Units | YTD $ | YTD Units | ITD $ | ITD Units
  | Commitments | Analytics | Lock-up & Redemption | Side Letter | Waterfall
"""

from __future__ import annotations

import re
import pandas as pd


# ── Normalisation ──────────────────────────────────────────────────────────────

def _norm(name: str) -> str:
    """Lowercase, strip, collapse all whitespace (including \\n from wrapped Excel cells)."""
    return re.sub(r"\s+", " ", str(name).replace("\n", " ")).strip().lower()


def _parse_num(val) -> float | None:
    """Parse formatted numeric strings: $1,234.56  (1,234)  1.02x  12%  -  —"""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ("-", "—", "", "nan", "n/a"):
        return 0.0
    s = s.replace(",", "").replace("$", "").replace("%", "").replace("x", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


# ── Column alias table ─────────────────────────────────────────────────────────
# (internal_name, exact_normalized_column_name)
# First match wins when the same normalized name appears twice (duplicates).

_ALIASES: list[tuple[str, str]] = [
    # ── Investor identifier ────────────────────────────────────────────────────
    ("INVESTOR_NAME",       "investor name - legal name from master list"),

    # ── Unit Prices ────────────────────────────────────────────────────────────
    ("BEG_PX",              "beginning partner's capital unit price"),
    ("XFER_IN_PX",          "transfer of units price in"),
    ("XFER_OUT_PX",         "transfer of units price out"),
    ("CONTRIB_PX",          "capital contribution unit price"),
    ("INC_PX",              "investment and other income unit price"),
    ("EXP_PX",              "fund level expense unit price"),
    ("UNRLZ_PX",            "net unrealized gain(loss) unit price"),
    ("RLZD_PX",             "net realized gain (loss) unit price"),
    ("EQ_PRED_PX",          "partner's equity before distributions unit price"),
    ("DIST_LP_PX",          "distr declared to lps unit price"),
    ("DIST_MGR_PX",         "distr redirected to mgr for fees unit price"),
    ("INC_FEE_PX",          "incentive fee unit price"),
    ("TAX_RED_PX",          "reduction of distributions for investor specific taxes unit price"),
    ("END_PX",              "ending partner's capital unit price"),

    # ── CQ Capital ────────────────────────────────────────────────────────────
    ("BEG_CAP_CQ",          "beginning partner's capital cq"),
    ("XFER_IN_CQ",          "transfer of units in cq"),
    ("XFER_OUT_CQ",         "transfer of units out cq"),
    ("CONTRIB_CQ",          "capital contribution cq"),
    ("DRIP_CQ",             "contr from div reinvest cq"),
    ("REDEMP_CQ",           "capital redemption cq"),
    ("INC_CQ",              "investment income (loss before fees) cq"),
    ("EXP_CQ",              "fund level expense cq"),
    ("UNRLZ_CQ",            "net unrealized gain(loss) cq"),
    ("RLZD_CQ",             "net realized gain (loss) cq"),
    ("EQ_PRED_CQ",          "partner's equity before distributions cq"),
    ("DIST_LP_CQ",          "distr declared to lps cq"),
    ("DIST_MGR_CQ",         "distri redirected to mgr for fees cq"),
    ("INC_FEE_CQ",          "incentive fee cq"),
    ("TAX_RED_CQ",          "reduction of distributions for investor specific taxes cq"),
    ("END_CAP_CQ",          "ending partner's capital cq"),

    # ── CQ Units ──────────────────────────────────────────────────────────────
    ("BEG_UNITS_CQ",        "beginning partner's capital units cq"),
    ("XFER_U_IN_CQ",        "transfer units in cq"),
    ("XFER_U_OUT_CQ",       "transfer units out cq"),
    ("CONTRIB_U_CQ",        "capital contribution units cq"),
    ("DRIP_U_CQ",           "contr from div reinvest units cq"),
    ("REDEMP_U_CQ",         "capital redemption units cq"),
    ("END_UNITS_CQ",        "ending partner's capital units cq"),

    # ── YTD Capital ───────────────────────────────────────────────────────────
    ("BEG_CAP_YTD",         "beginning partner's capital ytd"),
    ("XFER_IN_YTD",         "transfer of units in ytd"),
    ("XFER_OUT_YTD",        "transfer of units out ytd"),
    ("CONTRIB_YTD",         "capital contribution ytd"),
    ("DRIP_YTD",            "contr from div reinvest ytd"),
    ("REDEMP_YTD",          "capital redemption ytd"),
    ("INC_YTD",             "investment income (loss before fees) ytd"),
    ("EXP_YTD",             "fund level expense ytd"),
    ("UNRLZ_YTD",           "net unrealized gain(loss) ytd"),
    ("RLZD_YTD",            "net realized gain (loss) ytd"),
    ("EQ_PRED_YTD",         "partner's capital before dividends ytd"),
    ("DIST_LP_YTD",         "distr declared to lps ytd"),
    ("DIST_MGR_YTD",        "distr redirected to mgr for fees ytd"),
    ("INC_FEE_YTD",         "incentive fee ytd"),
    ("TAX_RED_YTD",         "reduction of distributions for investor specific taxes ytd"),
    ("END_CAP_YTD",         "ending partner's capital ytd"),

    # ── YTD Units ─────────────────────────────────────────────────────────────
    ("BEG_UNITS_YTD",       "beginning partner's capital units ytd"),
    ("XFER_U_IN_YTD",       "transfer units in ytd"),
    ("XFER_U_OUT_YTD",      "transfer units out ytd"),
    ("CONTRIB_U_YTD",       "capital contribution units ytd"),
    ("DRIP_U_YTD",          "contr from div reinvest units ytd"),
    ("REDEMP_U_YTD",        "capital redemption units ytd"),
    ("END_UNITS_YTD",       "ending partner's capital units ytd"),

    # ── ITD Capital ───────────────────────────────────────────────────────────
    ("BEG_CAP_ITD",         "beginning partner's capital itd"),
    ("XFER_IN_ITD",         "transfer of units in itd"),
    ("XFER_OUT_ITD",        "transfer of units out itd"),
    ("CONTRIB_ITD",         "capital contribution itd"),
    ("DRIP_ITD",            "contr from div reinvest itd"),
    ("REDEMP_ITD",          "capital redemption itd"),
    ("INC_ITD",             "investment income (loss before fees) itd"),
    ("EXP_ITD",             "fund level expense itd"),
    ("UNRLZ_ITD",           "net unrealized gain(loss) itd"),
    ("RLZD_ITD",            "net realized gain (loss) itd"),
    ("EQ_PRED_ITD",         "partner's capital before dividend itd"),
    ("DIST_LP_ITD",         "distr declared to lps itd"),
    ("DIST_MGR_ITD",        "distri redirected to mgr for fee itd"),   # singular "fee"
    ("INC_FEE_ITD",         "incentive fees itd"),                      # plural "fees"
    ("TAX_RED_ITD",         "reduction of distributions for investor specific taxes itd"),
    ("END_CAP_ITD",         "ending partner's capital itd"),

    # ── ITD Units ─────────────────────────────────────────────────────────────
    ("BEG_UNITS_ITD",       "beginning partner's capital units itd"),
    ("XFER_U_IN_ITD",       "transfer units in itd"),
    ("XFER_U_OUT_ITD",      "transfer units out itd"),
    ("CONTRIB_U_ITD",       "capital contribution units itd"),
    ("DRIP_U_ITD",          "contr from div reinvest units itd"),
    ("REDEMP_U_ITD",        "capital redemption units itd"),
    ("END_UNITS_ITD",       "ending partner's capital units itd"),

    # ── Commitments ───────────────────────────────────────────────────────────
    ("TOTAL_COMMIT",        "total capital commitment"),
    ("FUNDED_COMMIT",       "funded capital commitment"),
    ("XFER_COMMIT",         "transfer of commitment"),
    ("AVAIL_COMMIT",        "available commitment"),

    # ── Analytics ─────────────────────────────────────────────────────────────
    ("GROSS_IRR",           "gross irr"),
    ("NET_IRR",             "net irr"),
    ("INCEPTION_DATE",      "investor inception date"),
    ("PCT_AUM",             "% of fund aum"),
    ("DPI",                 "dpi"),
    ("RVPI",                "rvpi"),
    ("TVPI",                "tvpi"),
    ("COMMIT_FUNDED_PCT",   "commitment funded %"),
    ("UNRLZ_CQ_DLR",        "unrealized gain cq ($)"),
    ("RLZD_CQ_DLR",         "realized gain cq ($)"),
    ("TOT_RET_CQ_DLR",      "total return cq ($)"),
    ("TOT_RET_CQ_PCT",      "total return cq %"),
    ("UNRLZ_ITD_DLR",       "unrealized gain itd ($)"),
    ("RLZD_ITD_DLR",        "realized gain itd ($)"),
    ("TOT_RET_ITD_DLR",     "total return itd ($)"),
    ("TOT_RET_ITD_PCT",     "total return itd %"),
    ("MGMT_FEE_ITD_DLR",    "mgmt fee itd ($)"),
    ("INC_FEE_ITD_DLR",     "incentive fee itd ($)"),
    ("TOT_FEES_ITD_DLR",    "total fees itd ($)"),
    ("NET_RET_ITD_DLR",     "net return itd ($)"),
    ("INC_FEE_RATE",        "incentive fee rate"),   # first occurrence (Analytics)
    ("PREF_RET",            "preferred return"),      # first occurrence (Analytics)
    ("HURDLE_EXCEEDED",     "hurdle exceeded?"),

    # ── Lock-up & Redemption Terms ────────────────────────────────────────────
    ("LOCKUP_MO",           "lock-up period (months)"),
    ("SUB_DATE",            "subscription date"),
    ("FIRST_CONTRIB_DATE",  "first contribution date"),
    ("REDEMP_ELIG_DATE",    "redemption eligibility date"),
    ("LOCKUP_EXPIRED",      "lock-up expired?"),
    ("REDEMP_FREQ",         "redemption frequency"),
    ("GATE_PROV",           "gate provision"),
    ("NOTICE_DAYS",         "notice period (days)"),
    ("MONTHS_REM",          "months remaining"),
    ("SIDE_POCKET",         "side pocket eligible?"),
    ("DRIP_ENROLLED",       "drip enrolled?"),
    ("DIST_PREF",           "distribution preference"),
    ("HWM_ACTIVE",          "high-water mark active?"),
    ("REPT_CCY",            "reporting currency"),
    ("FATCA",               "fatca status"),
    ("AML_KYC",             "aml/kyc status"),
    ("ACCREDITED",          "accredited / qualified"),
    ("CUSTODIAN",           "custodian / prime broker"),
    ("SIDE_LETTER_FLG",     "side letter flag"),
    ("SPECIAL_TERMS",       "special terms notes"),

    # ── Side Letter Terms ─────────────────────────────────────────────────────
    ("MGMT_FEE_RATE",       "mgmt fee rate"),
    # "incentive fee rate" and "preferred return" are duplicates of Analytics cols;
    # second occurrences will be suffixed ".1" by pandas — handled in reader.
    ("HURDLE_TYPE",         "hurdle type"),
    ("CATCHUP_PCT",         "catch-up %"),

    # ── Waterfall & Fee Parameters ────────────────────────────────────────────
    ("HURDLE_AMT_ITD",      "hurdle amount itd ($)"),
    # "total return itd ($)" is a duplicate — second occurrence skipped
    ("EXCESS_HURDLE",       "excess over hurdle ($)"),
    ("GP_CATCHUP_AMT",      "gp catch-up amount ($)"),
    ("LP_NET_WF",           "lp net waterfall share ($)"),
    ("WF_TIER",             "waterfall tier"),
]

# Primary lookup: normalized_name → internal_name (first match wins)
_NORM_TO_INTERNAL: dict[str, str] = {}
for _int, _n in _ALIASES:
    if _n not in _NORM_TO_INTERNAL:
        _NORM_TO_INTERNAL[_n] = _int

# Text-valued columns (not parsed as numeric)
_TEXT_COLS = {
    "INVESTOR_NAME", "INCEPTION_DATE", "LOCKUP_EXPIRED", "REDEMP_FREQ",
    "GATE_PROV", "SIDE_POCKET", "DRIP_ENROLLED", "DIST_PREF", "HWM_ACTIVE",
    "REPT_CCY", "FATCA", "AML_KYC", "ACCREDITED", "CUSTODIAN",
    "SIDE_LETTER_FLG", "SPECIAL_TERMS", "HURDLE_EXCEEDED", "HURDLE_TYPE",
    "WF_TIER", "SUB_DATE", "FIRST_CONTRIB_DATE", "REDEMP_ELIG_DATE",
}

# Minimum columns required for statement generation
HF_REQUIRED_COLS = {
    "ending partner's capital cq",
    "investor name - legal name from master list",
}


# ── Smart header detection ─────────────────────────────────────────────────────

def _detect_header_row(df_raw: pd.DataFrame) -> int:
    """Return the 0-based row index that contains the real column headers."""
    for i, row in df_raw.iterrows():
        for val in row:
            if isinstance(val, str) and (
                "investor name" in val.lower()
                or "legal name" in val.lower()
                or "beginning partner" in val.lower()
            ):
                return int(i)
    return 0


def read_hf_pcap_from_upload(uploaded_file) -> pd.DataFrame:
    """
    Smart reader for the KPMG HF PCAP Excel.

    Handles the two-row header pattern:
      Row 1 — merged group headers (LOCK-UP & REDEMPTION TERMS, etc.)
      Row 2 — actual column names
      Row 3+ — investor data rows
    """
    # First pass: detect which row holds the real column headers
    try:
        raw = pd.read_excel(uploaded_file, header=None, nrows=5)
        hdr_row = _detect_header_row(raw)
    except Exception:
        hdr_row = 1

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=hdr_row)
    return read_hf_pcap(df)


# ── Main normalisation function ────────────────────────────────────────────────

def read_hf_pcap(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise a KPMG HF PCAP DataFrame.

    1. Maps column names to short internal names.
    2. Drops blank rows and the fund totals row (no investor name).
    3. Parses formatted numeric strings → float (or keeps as str for text cols).

    Returns a clean DataFrame with one row per investor.
    """
    df = df_raw.copy()

    # Build rename map; skip columns that can't be matched
    seen_norms: set[str] = set()
    rename_map: dict[str, str] = {}

    for orig_col in df.columns:
        normed = _norm(str(orig_col))
        if normed in seen_norms:
            continue  # duplicate — keep first occurrence only
        seen_norms.add(normed)
        if normed in _NORM_TO_INTERNAL:
            rename_map[orig_col] = _NORM_TO_INTERNAL[normed]

    df.rename(columns=rename_map, inplace=True)

    # Fallback: try to locate investor name column by content if not mapped
    if "INVESTOR_NAME" not in df.columns:
        for col in df.columns:
            sample = df[col].dropna().astype(str).head(10)
            keywords = ["LP", "LLC", "Fund", "Trust", "Office", "Partners", "Capital"]
            if any(any(kw in v for kw in keywords) for v in sample):
                df.rename(columns={col: "INVESTOR_NAME"}, inplace=True)
                break

    if "INVESTOR_NAME" not in df.columns:
        raise ValueError(
            "Cannot identify investor name column. "
            "Expected 'Investor Name - Legal Name from Master List'."
        )

    # Drop rows with no investor name (blank rows, totals row)
    df["INVESTOR_NAME"] = df["INVESTOR_NAME"].astype(str).str.strip()
    df = df[
        df["INVESTOR_NAME"].notna()
        & ~df["INVESTOR_NAME"].isin(["", "nan", "None", "NaN"])
    ].copy()

    # Parse numeric columns
    for col in df.columns:
        if col in _TEXT_COLS:
            df[col] = df[col].astype(str).str.strip().replace("nan", "").replace("None", "")
        elif df[col].dtype == object:
            df[col] = df[col].apply(_parse_num)

    return df.reset_index(drop=True)


# ── Convenience accessor ───────────────────────────────────────────────────────

def _get(row: "pd.Series", field: str, default=0.0):
    """Safe field accessor — returns default if field absent or NaN."""
    val = row.get(field, default)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    return val


def _gets(row: "pd.Series", field: str, default: str = "—") -> str:
    """Safe string accessor."""
    val = row.get(field, default)
    s = str(val).strip() if val is not None else ""
    return s if s and s not in ("nan", "None", "") else default
