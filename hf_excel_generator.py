"""
Hedge Fund PCAP — 10-sheet openpyxl workbook generator.

Input:
  pcap_df      — clean DataFrame from hf_pcap_engine.read_hf_pcap()
  cf_ledger_df — optional transaction-level cashflow DataFrame

Sheets:
  1. Period_Params          — fund parameters (single source of truth)
  2. Dashboard              — fund KPI summary + investor breakdown
  3. Investor_Register      — per-investor static / compliance data
  4. PCAP                   — full raw PCAP (all columns)
  5. Capital_Accounts       — simplified 9-line per-investor CQ/YTD/ITD
  6. CF_Ledger              — transaction-level cashflow (optional input)
  7. CF_Aggregator          — investor × quarter aggregation
  8. Distribution_Waterfall — waterfall analysis per investor
  9. Cashflow_IRR           — IRR cashflow vectors + XIRR formula
 10. Stress_IRR             — NAV haircut scenarios (0%, −10%, −20%, −30%, custom)
"""

from __future__ import annotations

import io
import math
from datetime import date, datetime
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ── KPMG colour palette ────────────────────────────────────────────────────────
_BLUE  = "00338D"
_LIGHT = "ACEAFF"
_NAVY  = "0C233C"
_WHITE = "FFFFFF"
_GRAY1 = "F2F6FF"
_GRAY2 = "FFFFFF"
_TOTAL = "E8F4FF"
_SEC   = "D0E8FF"
_WARN  = "FFF2CC"


# ── Style factories ────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(bold=False, size=9, color=_NAVY, italic=False) -> Font:
    return Font(name="Arial", bold=bold, size=size, color=color, italic=italic)


def _align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _border(color=_LIGHT) -> Border:
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _hdr(ws, row: int, col: int, value: str, width: float | None = None) -> None:
    c = ws.cell(row=row, column=col, value=value)
    c.font      = _font(bold=True, color=_WHITE, size=9)
    c.fill      = _fill(_BLUE)
    c.alignment = _align("center", wrap=True)
    c.border    = _border()
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width


def _cell(ws, row: int, col: int, value: Any = None,
          fmt: str | None = None, bold: bool = False,
          align: str = "right", fill_hex: str | None = None) -> None:
    c = ws.cell(row=row, column=col, value=value)
    c.font      = _font(bold=bold, size=9)
    c.alignment = _align(align)
    c.border    = _border()
    if fmt:
        c.number_format = fmt
    if fill_hex:
        c.fill = _fill(fill_hex)


def _banner(ws, title: str, subtitle: str = "", n_cols: int = 12) -> int:
    last = get_column_letter(n_cols)
    ws.merge_cells(f"A1:{last}1")
    c = ws["A1"]
    c.value     = title
    c.font      = _font(bold=True, size=13, color=_WHITE)
    c.fill      = _fill(_BLUE)
    c.alignment = _align("center")
    ws.row_dimensions[1].height = 26
    if subtitle:
        ws.merge_cells(f"A2:{last}2")
        c2 = ws["A2"]
        c2.value     = subtitle
        c2.font      = _font(size=9, italic=True)
        c2.fill      = _fill(_LIGHT)
        c2.alignment = _align("center")
        return 3
    return 2


def _section_banner(ws, row: int, text: str, n_cols: int) -> None:
    end_col = get_column_letter(n_cols)
    ws.merge_cells(f"A{row}:{end_col}{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.font      = _font(bold=True, color=_NAVY, size=10)
    c.fill      = _fill(_SEC)
    c.alignment = _align("left")
    ws.row_dimensions[row].height = 16


def _inv_banner(ws, row: int, investor: str, n_cols: int) -> None:
    end_col = get_column_letter(n_cols)
    ws.merge_cells(f"A{row}:{end_col}{row}")
    c = ws.cell(row=row, column=1, value=f"Investor: {investor}")
    c.font      = _font(bold=True, color=_WHITE, size=10)
    c.fill      = _fill(_NAVY)
    c.alignment = _align("left")
    ws.row_dimensions[row].height = 18


def _placeholder(ws, row: int, message: str, n_cols: int = 8) -> None:
    end_col = get_column_letter(n_cols)
    ws.merge_cells(f"A{row}:{end_col}{row}")
    c = ws.cell(row=row, column=1, value=message)
    c.font      = _font(italic=True, color="888888", size=10)
    c.fill      = _fill(_WARN)
    c.alignment = _align("center")


# ── Number formats ─────────────────────────────────────────────────────────────
_USD   = '"$"#,##0.00'
_PCT_R = '0.00"%"'
_UNITS = '#,##0.0000'
_PX    = '"$"#,##0.0000'
_X     = '0.00"x"'
_INT   = '#,##0'
_DATEFMT = 'YYYY-MM-DD'


# ── Safe accessors ─────────────────────────────────────────────────────────────

def _gv(row: pd.Series, field: str, default=None):
    val = row.get(field, default)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    return val


def _gf(row: pd.Series, field: str, default: float = 0.0) -> float:
    val = _gv(row, field, default)
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _gt(row: pd.Series, field: str, default: str = "—") -> str:
    val = _gv(row, field, "")
    s = str(val).strip() if val is not None else ""
    return s if s and s not in ("nan", "None", "") else default


def _col_sum(pcap: pd.DataFrame, col: str) -> float:
    if col in pcap.columns:
        return float(pcap[col].dropna().sum())
    return 0.0


def _col_mean(pcap: pd.DataFrame, col: str) -> float:
    if col in pcap.columns:
        s = pcap[col].dropna()
        return float(s.mean()) if len(s) > 0 else 0.0
    return 0.0


def _col_first(pcap: pd.DataFrame, col: str, default=None):
    if col in pcap.columns:
        vals = pcap[col].dropna()
        if len(vals) > 0:
            return vals.iloc[0]
    return default


def _parse_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if v is None:
        return None
    try:
        return datetime.strptime(str(v).strip()[:10], "%Y-%m-%d").date()
    except Exception:
        pass
    try:
        return datetime.strptime(str(v).strip(), "%m/%d/%Y").date()
    except Exception:
        return None


def _stress_irr(stressed_nav: float, dist_itd: float, contrib_itd: float,
                hold_days: int) -> float:
    if contrib_itd <= 0 or hold_days <= 0:
        return 0.0
    terminal = stressed_nav + dist_itd
    if terminal <= 0:
        return -1.0
    try:
        return (terminal / contrib_itd) ** (365.0 / hold_days) - 1.0
    except Exception:
        return 0.0


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 1: Period_Params
# ═══════════════════════════════════════════════════════════════════════════════

def _build_period_params(ws, pcap: pd.DataFrame) -> None:
    ws.title = "Period_Params"
    r = _banner(ws, "Period Parameters — Single Source of Truth",
                "Fund-level parameters derived from pre-calculated PCAP data", n_cols=6)

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 26

    _section_banner(ws, r, "A.  Fund-Level Parameters", 3)
    r += 1
    _hdr(ws, r, 1, "Parameter")
    _hdr(ws, r, 2, "Value")
    _hdr(ws, r, 3, "Notes")
    r += 1

    beg_px  = _col_first(pcap, "BEG_PX",  0.0)
    end_px  = _col_first(pcap, "END_PX",  0.0)

    params: list[tuple] = [
        ("Number of Investors",           int(len(pcap)),                      None,    "Integer count of LP investors"),
        ("Total AUM — Ending NAV (CQ)",   _col_sum(pcap, "END_CAP_CQ"),        _USD,    "Sum of all ending partner's capital (CQ)"),
        ("Total AUM — Ending NAV (ITD)",  _col_sum(pcap, "END_CAP_ITD"),       _USD,    "Sum of all ending partner's capital (ITD)"),
        ("Total Contributions (ITD)",     _col_sum(pcap, "CONTRIB_ITD"),        _USD,    "Sum of all LP capital contributions"),
        ("Total LP Distributions (ITD)",  _col_sum(pcap, "DIST_LP_ITD"),       _USD,    "Sum of all distributions to LPs"),
        ("Total Incentive Fees (ITD)",    _col_sum(pcap, "INC_FEE_ITD"),        _USD,    "Sum of all incentive/performance fees"),
        ("Total Mgmt. Fees (ITD)",        _col_sum(pcap, "MGMT_FEE_ITD_DLR"),  _USD,    "Sum of all management fees (dollar)"),
        ("Opening Unit Price",            float(beg_px) if beg_px else 0.0,    _PX,     "Beginning of period NAV per unit"),
        ("Ending Unit Price",             float(end_px) if end_px else 0.0,    _PX,     "End of period NAV per unit"),
        ("Unit Price Change",             (float(end_px) - float(beg_px)) if (beg_px and end_px) else 0.0, _PX, "Ending − Opening unit price"),
        ("Average Gross IRR",             _col_mean(pcap, "GROSS_IRR"),         _PCT_R,  "Equal-weighted average across all investors"),
        ("Average Net IRR",               _col_mean(pcap, "NET_IRR"),           _PCT_R,  "Equal-weighted average across all investors"),
        ("Average TVPI",                  _col_mean(pcap, "TVPI"),              _X,      "Equal-weighted average"),
        ("Average DPI",                   _col_mean(pcap, "DPI"),               _X,      "Equal-weighted average"),
        ("Average RVPI",                  _col_mean(pcap, "RVPI"),              _X,      "Equal-weighted average"),
        ("LP Net Waterfall (Total ITD)",  _col_sum(pcap, "LP_NET_WF"),          _USD,    "Sum of all LP net waterfall shares"),
    ]

    for i, (label, val, fmt, notes) in enumerate(params):
        fill = _GRAY1 if i % 2 else _GRAY2
        row_idx = r + i
        _cell(ws, row_idx, 1, label, align="left",  fill_hex=fill)
        _cell(ws, row_idx, 2, val,   fmt=fmt, align="right" if fmt else "left", fill_hex=fill)
        _cell(ws, row_idx, 3, notes, align="left",  fill_hex=fill)
    r += len(params)

    # Fee structures
    r += 1
    _section_banner(ws, r, "B.  Fee Structures by Investor", 6)
    r += 1
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14
    _hdr(ws, r, 1, "Investor",       width=36)
    _hdr(ws, r, 2, "Mgmt. Fee %",    width=16)
    _hdr(ws, r, 3, "Incentive Fee %",width=18)
    _hdr(ws, r, 4, "Pref. Return %", width=16)
    _hdr(ws, r, 5, "Hurdle Type",    width=16)
    _hdr(ws, r, 6, "Catch-up %",     width=14)
    r += 1

    for i, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        _cell(ws, r, 1, _gt(row, "INVESTOR_NAME"), align="left",  fill_hex=fill)
        _cell(ws, r, 2, _gf(row, "MGMT_FEE_RATE"), fmt=_PCT_R,   fill_hex=fill)
        _cell(ws, r, 3, _gf(row, "INC_FEE_RATE"),  fmt=_PCT_R,   fill_hex=fill)
        _cell(ws, r, 4, _gf(row, "PREF_RET"),      fmt=_PCT_R,   fill_hex=fill)
        _cell(ws, r, 5, _gt(row, "HURDLE_TYPE"),   align="left",  fill_hex=fill)
        _cell(ws, r, 6, _gf(row, "CATCHUP_PCT"),   fmt=_PCT_R,   fill_hex=fill)
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 2: Dashboard
# ═══════════════════════════════════════════════════════════════════════════════

def _build_dashboard(ws, pcap: pd.DataFrame) -> None:
    ws.title = "Dashboard"
    r = _banner(ws, "Fund Dashboard — KPI Summary", n_cols=9)

    kpis = [
        ("Total Investors",               int(len(pcap)),                   None,   "left"),
        ("Total AUM (Ending NAV, CQ)",    _col_sum(pcap, "END_CAP_CQ"),     _USD,   "right"),
        ("Total Contributions (ITD)",     _col_sum(pcap, "CONTRIB_ITD"),    _USD,   "right"),
        ("Total LP Distributions (ITD)",  _col_sum(pcap, "DIST_LP_ITD"),   _USD,   "right"),
        ("Total Incentive Fees (ITD)",    _col_sum(pcap, "INC_FEE_ITD"),   _USD,   "right"),
        ("LP Net Waterfall (ITD)",        _col_sum(pcap, "LP_NET_WF"),     _USD,   "right"),
        ("Average Gross IRR",             _col_mean(pcap, "GROSS_IRR"),    _PCT_R, "right"),
        ("Average Net IRR",               _col_mean(pcap, "NET_IRR"),      _PCT_R, "right"),
        ("Average TVPI",                  _col_mean(pcap, "TVPI"),         _X,     "right"),
        ("Ending Unit Price",             _col_first(pcap, "END_PX", 0.0), _PX,    "right"),
    ]

    _hdr(ws, r, 1, "KPI",   width=34)
    _hdr(ws, r, 2, "Value", width=22)
    r += 1

    for i, (label, val, fmt, align) in enumerate(kpis):
        fill = _GRAY1 if i % 2 else _GRAY2
        _cell(ws, r, 1, label, align="left",  fill_hex=fill)
        _cell(ws, r, 2, val,   fmt=fmt, align=align, fill_hex=fill)
        r += 1

    # Investor breakdown table
    r += 1
    _section_banner(ws, r, "Investor Breakdown", 9)
    r += 1

    hdrs   = ["Investor", "Inception", "Currency", "NAV (CQ)", "NAV (ITD)",
              "Gross IRR", "Net IRR", "TVPI", "LP Net WF"]
    widths = [26, 14, 10, 18, 18, 12, 12, 10, 18]
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        _hdr(ws, r, ci, h, width=w)
    r += 1

    for i, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        vals = [
            (_gt(row, "INVESTOR_NAME"),  "left",  None),
            (_gt(row, "INCEPTION_DATE"), "left",  None),
            (_gt(row, "REPT_CCY"),       "left",  None),
            (_gf(row, "END_CAP_CQ"),    "right", _USD),
            (_gf(row, "END_CAP_ITD"),   "right", _USD),
            (_gf(row, "GROSS_IRR"),     "right", _PCT_R),
            (_gf(row, "NET_IRR"),       "right", _PCT_R),
            (_gf(row, "TVPI"),          "right", _X),
            (_gf(row, "LP_NET_WF"),     "right", _USD),
        ]
        for ci, (v, align, fmt) in enumerate(vals, 1):
            _cell(ws, r, ci, v, fmt=fmt, align=align, fill_hex=fill)
        r += 1

    # Totals row
    fill = _TOTAL
    _cell(ws, r, 1, "TOTAL / AVERAGE", bold=True, align="left", fill_hex=fill)
    _cell(ws, r, 2, "",   align="left",  fill_hex=fill)
    _cell(ws, r, 3, "",   align="left",  fill_hex=fill)
    _cell(ws, r, 4, _col_sum(pcap, "END_CAP_CQ"),  fmt=_USD,   bold=True, fill_hex=fill)
    _cell(ws, r, 5, _col_sum(pcap, "END_CAP_ITD"), fmt=_USD,   bold=True, fill_hex=fill)
    _cell(ws, r, 6, _col_mean(pcap, "GROSS_IRR"),  fmt=_PCT_R, bold=True, fill_hex=fill)
    _cell(ws, r, 7, _col_mean(pcap, "NET_IRR"),    fmt=_PCT_R, bold=True, fill_hex=fill)
    _cell(ws, r, 8, _col_mean(pcap, "TVPI"),       fmt=_X,     bold=True, fill_hex=fill)
    _cell(ws, r, 9, _col_sum(pcap, "LP_NET_WF"),   fmt=_USD,   bold=True, fill_hex=fill)


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 3: Investor_Register
# ═══════════════════════════════════════════════════════════════════════════════

def _build_investor_register(ws, pcap: pd.DataFrame) -> None:
    ws.title = "Investor_Register"
    r = _banner(ws, "Investor Register — Static & Compliance Data", n_cols=16)

    cols_a = [
        ("Investor Name",       "INVESTOR_NAME",       "left",  None,   24),
        ("Entity Type",         "ENTITY_TYPE",         "left",  None,   18),
        ("Tax ID",              "TAX_ID",              "left",  None,   16),
        ("Jurisdiction",        "JURISDICTION",        "left",  None,   16),
        ("Inception Date",      "INCEPTION_DATE",      "left",  None,   14),
        ("Subscription Date",   "SUB_DATE",            "left",  None,   14),
        ("1st Contribution",    "FIRST_CONTRIB_DATE",  "left",  None,   14),
        ("Currency",            "REPT_CCY",            "left",  None,   10),
        ("Mgmt. Fee %",         "MGMT_FEE_RATE",       "right", _PCT_R, 12),
        ("Incentive Fee %",     "INC_FEE_RATE",        "right", _PCT_R, 14),
        ("Pref. Return %",      "PREF_RET",            "right", _PCT_R, 14),
        ("Hurdle Type",         "HURDLE_TYPE",         "left",  None,   14),
        ("Catch-up %",          "CATCHUP_PCT",         "right", _PCT_R, 12),
        ("Side Letter",         "SIDE_LETTER_FLG",     "left",  None,   12),
        ("FATCA",               "FATCA",               "left",  None,   14),
        ("AML/KYC",             "AML_KYC",             "left",  None,   14),
        ("Accredited",          "ACCREDITED",          "left",  None,   14),
        ("Custodian",           "CUSTODIAN",           "left",  None,   22),
        ("Special Terms",       "SPECIAL_TERMS",       "left",  None,   22),
        ("Lock-up (months)",    "LOCKUP_MO",           "right", _INT,   14),
        ("Lock-up Expired?",    "LOCKUP_EXPIRED",      "left",  None,   14),
        ("Redemp. Frequency",   "REDEMP_FREQ",         "left",  None,   16),
        ("Gate Provision",      "GATE_PROV",           "left",  None,   18),
        ("Notice (days)",       "NOTICE_DAYS",         "right", _INT,   13),
        ("DRIP Enrolled?",      "DRIP_ENROLLED",       "left",  None,   12),
        ("Side Pocket?",        "SIDE_POCKET",         "left",  None,   12),
    ]

    for ci, (label, _, align, fmt, width) in enumerate(cols_a, 1):
        _hdr(ws, r, ci, label, width=width)
    r += 1

    for i, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        for ci, (_, field, align, fmt, _w) in enumerate(cols_a, 1):
            try:
                v = float(_gv(row, field, 0.0)) if fmt in (_PCT_R, _INT) else _gt(row, field)
            except Exception:
                v = _gt(row, field)
            _cell(ws, r, ci, v, fmt=fmt, align=align, fill_hex=fill)
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 4: PCAP
# ═══════════════════════════════════════════════════════════════════════════════

def _build_pcap(ws, pcap: pd.DataFrame) -> None:
    ws.title = "PCAP"
    r = _banner(ws, "PCAP — Full Pre-Calculated Data",
                "Complete investor PCAP as loaded from source Excel (all internal column names)", n_cols=16)

    cols = list(pcap.columns)
    for ci, col in enumerate(cols, 1):
        _hdr(ws, r, ci, col, width=16)
    r += 1

    for ri, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if ri % 2 else _GRAY2
        for ci, col in enumerate(cols, 1):
            val = _gv(row, col)
            if isinstance(val, float):
                _cell(ws, r, ci, val, fill_hex=fill)
            else:
                _cell(ws, r, ci, str(val) if val is not None else "",
                      align="left", fill_hex=fill)
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 5: Capital_Accounts
# ═══════════════════════════════════════════════════════════════════════════════

def _build_capital_accounts(ws, pcap: pd.DataFrame) -> None:
    ws.title = "Capital_Accounts"
    r = _banner(ws, "Capital Accounts — 3-Period View (CQ / YTD / ITD)", n_cols=4)

    _hdr(ws, r, 1, "Line Item",  width=38)
    _hdr(ws, r, 2, "CQ ($)",     width=20)
    _hdr(ws, r, 3, "YTD ($)",    width=20)
    _hdr(ws, r, 4, "ITD ($)",    width=20)
    r += 1

    lines = [
        ("Beginning Partner's Capital",   "BEG_CAP_CQ",  "BEG_CAP_YTD",  "BEG_CAP_ITD",  False),
        ("(+) Capital Contributions",     "CONTRIB_CQ",  "CONTRIB_YTD",  "CONTRIB_ITD",  False),
        ("(+) DRIP Reinvestment",         "DRIP_CQ",     "DRIP_YTD",     "DRIP_ITD",     False),
        ("(+) Investment Income (Loss)",  "INC_CQ",      "INC_YTD",      "INC_ITD",      False),
        ("(+) Net Unrealized Gain (Loss)","UNRLZ_CQ",    "UNRLZ_YTD",    "UNRLZ_ITD",    False),
        ("(+) Net Realized Gain (Loss)",  "RLZD_CQ",     "RLZD_YTD",     "RLZD_ITD",     False),
        ("(–) Distributions to LP",       "DIST_LP_CQ",  "DIST_LP_YTD",  "DIST_LP_ITD",  False),
        ("(–) Incentive Fees",            "INC_FEE_CQ",  "INC_FEE_YTD",  "INC_FEE_ITD",  False),
        ("Ending Partner's Capital",      "END_CAP_CQ",  "END_CAP_YTD",  "END_CAP_ITD",  True),
    ]

    for _, row in pcap.iterrows():
        _inv_banner(ws, r, _gt(row, "INVESTOR_NAME"), 4)
        r += 1
        for li, (label, cq_f, ytd_f, itd_f, bold) in enumerate(lines):
            fill = _TOTAL if bold else (_GRAY1 if li % 2 else _GRAY2)
            _cell(ws, r, 1, label,               align="left", bold=bold, fill_hex=fill)
            _cell(ws, r, 2, _gf(row, cq_f),  fmt=_USD, bold=bold, fill_hex=fill)
            _cell(ws, r, 3, _gf(row, ytd_f), fmt=_USD, bold=bold, fill_hex=fill)
            _cell(ws, r, 4, _gf(row, itd_f), fmt=_USD, bold=bold, fill_hex=fill)
            r += 1
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 6: CF_Ledger
# ═══════════════════════════════════════════════════════════════════════════════

def _build_cf_ledger(ws, cf_ledger: pd.DataFrame | None) -> None:
    ws.title = "CF_Ledger"
    r = _banner(ws, "Cashflow Ledger — Transaction-Level Data",
                "Upload a CF Ledger to populate. Required: INVESTOR_NAME, TRANSACTION_DATE, TYPE, AMOUNT", n_cols=9)

    if cf_ledger is None or cf_ledger.empty:
        _placeholder(ws, r, "No CF Ledger uploaded — provide a file with columns: "
                    "TRANSACTION_ID | INVESTOR_NAME | TRANSACTION_DATE | TYPE | AMOUNT | QUARTER | SUB_TYPE | UNITS | NOTES", 9)
        r += 1
        tmpl = ["TRANSACTION_ID", "INVESTOR_NAME", "TRANSACTION_DATE", "TYPE",
                "AMOUNT", "QUARTER", "SUB_TYPE", "UNITS", "NOTES"]
        for ci, col in enumerate(tmpl, 1):
            _hdr(ws, r, ci, col, width=18)
        return

    df = cf_ledger.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    for ci, col in enumerate(df.columns, 1):
        _hdr(ws, r, ci, col, width=18)
    r += 1

    for i, (_, row) in enumerate(df.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        for ci, col in enumerate(df.columns, 1):
            val = row[col]
            if isinstance(val, float) and pd.isna(val):
                val = ""
            _cell(ws, r, ci, val, align="left" if ci <= 4 else "right", fill_hex=fill)
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 7: CF_Aggregator
# ═══════════════════════════════════════════════════════════════════════════════

def _build_cf_aggregator(ws, pcap: pd.DataFrame, cf_ledger: pd.DataFrame | None) -> None:
    ws.title = "CF_Aggregator"
    r = _banner(ws, "CF Aggregator — Investor × Quarter Pivot",
                "Cashflow aggregation by investor and quarter", n_cols=10)

    if cf_ledger is None or cf_ledger.empty:
        _placeholder(ws, r, "Upload CF Ledger to populate this sheet", 10)
        return

    df = cf_ledger.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    required = {"INVESTOR_NAME", "TYPE", "AMOUNT"}
    if not required.issubset(set(df.columns)):
        _placeholder(ws, r, f"CF Ledger missing required columns: {required - set(df.columns)}", 10)
        return

    df["AMOUNT"] = pd.to_numeric(df["AMOUNT"], errors="coerce").fillna(0.0)

    if "QUARTER" not in df.columns:
        if "TRANSACTION_DATE" in df.columns:
            df["TRANSACTION_DATE"] = pd.to_datetime(df["TRANSACTION_DATE"], errors="coerce")
            df["QUARTER"] = df["TRANSACTION_DATE"].dt.to_period("Q").astype(str)
        else:
            df["QUARTER"] = "Unknown"

    records = []
    for (investor, quarter), grp in df.groupby(["INVESTOR_NAME", "QUARTER"], sort=True):
        contribs = grp.loc[grp["TYPE"] == "Contribution",  "AMOUNT"].sum()
        distribs = grp.loc[grp["TYPE"] == "Distribution",  "AMOUNT"].sum()
        drip     = grp.loc[grp["TYPE"] == "DRIP",          "AMOUNT"].sum()
        redemp   = grp.loc[grp["TYPE"] == "Redemption",    "AMOUNT"].sum()

        if "SUB_TYPE" in grp.columns:
            xfer_mask = grp["TYPE"] == "Transfer"
            xfer_in   = grp.loc[xfer_mask & (grp["SUB_TYPE"].astype(str).str.strip() == "In"),  "AMOUNT"].sum()
            xfer_out  = grp.loc[xfer_mask & (grp["SUB_TYPE"].astype(str).str.strip() == "Out"), "AMOUNT"].sum()
        else:
            xfer_in = xfer_out = 0.0

        net_cf = contribs - distribs - drip - redemp + xfer_in - xfer_out
        records.append({
            "INVESTOR_NAME": investor,
            "QUARTER":       quarter,
            "Contributions": contribs,
            "Distributions": distribs,
            "DRIP":          drip,
            "Transfers In":  xfer_in,
            "Transfers Out": xfer_out,
            "Redemptions":   redemp,
            "Net CF":        net_cf,
        })

    if not records:
        _placeholder(ws, r, "No aggregatable records in CF Ledger", 10)
        return

    hdrs   = ["Investor", "Quarter", "Contributions", "Distributions", "DRIP",
              "Transfers In", "Transfers Out", "Redemptions", "Net CF", "Running Balance"]
    widths = [26, 12, 16, 16, 14, 14, 14, 14, 16, 18]
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        _hdr(ws, r, ci, h, width=w)
    r += 1

    running: dict[str, float] = {}
    for i, rec in enumerate(records):
        inv = rec["INVESTOR_NAME"]
        running[inv] = running.get(inv, 0.0) + rec["Net CF"]
        fill = _GRAY1 if i % 2 else _GRAY2
        _cell(ws, r, 1,  rec["INVESTOR_NAME"], align="left", fill_hex=fill)
        _cell(ws, r, 2,  rec["QUARTER"],        align="left", fill_hex=fill)
        _cell(ws, r, 3,  rec["Contributions"],  fmt=_USD, fill_hex=fill)
        _cell(ws, r, 4,  rec["Distributions"],  fmt=_USD, fill_hex=fill)
        _cell(ws, r, 5,  rec["DRIP"],           fmt=_USD, fill_hex=fill)
        _cell(ws, r, 6,  rec["Transfers In"],   fmt=_USD, fill_hex=fill)
        _cell(ws, r, 7,  rec["Transfers Out"],  fmt=_USD, fill_hex=fill)
        _cell(ws, r, 8,  rec["Redemptions"],    fmt=_USD, fill_hex=fill)
        _cell(ws, r, 9,  rec["Net CF"],         fmt=_USD, fill_hex=fill)
        _cell(ws, r, 10, running[inv],          fmt=_USD, fill_hex=fill)
        r += 1

    # Totals
    _cell(ws, r, 1, "TOTAL", bold=True, align="left", fill_hex=_TOTAL)
    _cell(ws, r, 2, "",      align="left", fill_hex=_TOTAL)
    for ci, key in [(3,"Contributions"),(4,"Distributions"),(5,"DRIP"),
                    (6,"Transfers In"),(7,"Transfers Out"),(8,"Redemptions"),(9,"Net CF")]:
        total = sum(rec[key] for rec in records)
        _cell(ws, r, ci, total, fmt=_USD, bold=True, fill_hex=_TOTAL)


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 8: Distribution_Waterfall
# ═══════════════════════════════════════════════════════════════════════════════

def _build_distribution_waterfall(ws, pcap: pd.DataFrame) -> None:
    ws.title = "Distribution_Waterfall"
    r = _banner(ws, "Distribution Waterfall — ITD Per Investor",
                "Carry / waterfall mechanics: Total Return → Hurdle → Excess → GP Catch-up → LP Net", n_cols=9)

    hdrs   = ["Investor", "Contrib. ITD", "Pref. Ret %", "Hurdle Amt",
              "Total Return ITD", "Excess > Hurdle", "GP Catch-up", "LP Net WF", "WF Tier"]
    widths = [26, 18, 13, 18, 18, 18, 16, 18, 14]
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        _hdr(ws, r, ci, h, width=w)
    r += 1

    totals = {k: 0.0 for k in ["contrib","hurdle","tot_ret","excess","gp","lp"]}

    for i, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2

        contrib_itd = _gf(row, "CONTRIB_ITD")
        pref_ret    = _gf(row, "PREF_RET")
        end_cap_cq  = _gf(row, "END_CAP_CQ")
        dist_lp_itd = _gf(row, "DIST_LP_ITD")
        inc_fee_r   = _gf(row, "INC_FEE_RATE")
        catchup_pct = _gf(row, "CATCHUP_PCT", 100.0)

        # Prefer pre-computed values from PCAP; fall back to derived
        hurdle_amt  = _gf(row, "HURDLE_AMT_ITD") or (contrib_itd * pref_ret / 100.0)
        tot_ret_itd = _gf(row, "TOT_RET_ITD_DLR") or (end_cap_cq + dist_lp_itd - contrib_itd)
        excess      = _gf(row, "EXCESS_HURDLE")  if _gv(row, "EXCESS_HURDLE") is not None \
                      else max(tot_ret_itd - hurdle_amt, 0.0)
        gp_catchup  = _gf(row, "GP_CATCHUP_AMT") if _gv(row, "GP_CATCHUP_AMT") is not None \
                      else (inc_fee_r / 100.0) * (catchup_pct / 100.0) * excess
        lp_net      = _gf(row, "LP_NET_WF")      if _gv(row, "LP_NET_WF") is not None \
                      else max(excess - gp_catchup, 0.0)

        _cell(ws, r, 1, _gt(row, "INVESTOR_NAME"), align="left", fill_hex=fill)
        _cell(ws, r, 2, contrib_itd,  fmt=_USD,   fill_hex=fill)
        _cell(ws, r, 3, pref_ret,     fmt=_PCT_R, fill_hex=fill)
        _cell(ws, r, 4, hurdle_amt,   fmt=_USD,   fill_hex=fill)
        _cell(ws, r, 5, tot_ret_itd,  fmt=_USD,   fill_hex=fill)
        _cell(ws, r, 6, excess,       fmt=_USD,   fill_hex=fill)
        _cell(ws, r, 7, gp_catchup,   fmt=_USD,   fill_hex=fill)
        _cell(ws, r, 8, lp_net,       fmt=_USD,   bold=True, fill_hex=fill)
        _cell(ws, r, 9, _gt(row, "WF_TIER"), align="left", fill_hex=fill)
        r += 1

        totals["contrib"]  += contrib_itd
        totals["hurdle"]   += hurdle_amt
        totals["tot_ret"]  += tot_ret_itd
        totals["excess"]   += excess
        totals["gp"]       += gp_catchup
        totals["lp"]       += lp_net

    _cell(ws, r, 1, "TOTAL", bold=True, align="left", fill_hex=_TOTAL)
    _cell(ws, r, 2, totals["contrib"],  fmt=_USD, bold=True, fill_hex=_TOTAL)
    _cell(ws, r, 3, "",   align="left", fill_hex=_TOTAL)
    _cell(ws, r, 4, totals["hurdle"],   fmt=_USD, bold=True, fill_hex=_TOTAL)
    _cell(ws, r, 5, totals["tot_ret"],  fmt=_USD, bold=True, fill_hex=_TOTAL)
    _cell(ws, r, 6, totals["excess"],   fmt=_USD, bold=True, fill_hex=_TOTAL)
    _cell(ws, r, 7, totals["gp"],       fmt=_USD, bold=True, fill_hex=_TOTAL)
    _cell(ws, r, 8, totals["lp"],       fmt=_USD, bold=True, fill_hex=_TOTAL)


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 9: Cashflow_IRR
# ═══════════════════════════════════════════════════════════════════════════════

def _build_cashflow_irr(ws, pcap: pd.DataFrame, cf_ledger: pd.DataFrame | None) -> None:
    ws.title = "Cashflow_IRR"
    r = _banner(ws, "Cashflow IRR Analysis",
                "Gross & Net IRR cashflow vectors with XIRR formula. Values: negative=outflow, positive=inflow.", n_cols=8)

    if cf_ledger is not None and not cf_ledger.empty:
        _irr_from_ledger(ws, pcap, cf_ledger, r)
    else:
        _irr_from_pcap(ws, pcap, r)


def _irr_from_pcap(ws, pcap: pd.DataFrame, r: int) -> None:
    """2-date XIRR approximation using inception date and today."""
    today = date.today()

    # Layout:
    # Col A: Investor  B: CF_0 (inception)  C: CF_1 (terminal)  D: Date_0  E: Date_1
    # Col F: =XIRR(B:C, D:E) [Gross]   Col G: Net CF_1   Col H: =XIRR(B,G,D,E) [Net]
    # Col I: PCAP Gross IRR (check)     Col J: PCAP Net IRR (check)

    # ── Section A: Gross IRR ──────────────────────────────────────────────────
    _section_banner(ws, r, "A.  GROSS IRR — Pre-Fee Cashflows (2-date approximation)", 10)
    r += 1

    hdrs = ["Investor", "CF at Inception (−)", "CF at Terminal (+)",
            "Inception Date", "Terminal Date",
            "Gross XIRR (formula)", "PCAP Gross IRR (check)"]
    widths = [26, 22, 22, 16, 16, 22, 22]
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        _hdr(ws, r, ci, h, width=w)

    gross_data_start = r + 1

    for i, (_, row) in enumerate(pcap.iterrows()):
        r += 1
        fill = _GRAY1 if i % 2 else _GRAY2
        contrib_itd = _gf(row, "CONTRIB_ITD")
        end_cap_cq  = _gf(row, "END_CAP_CQ")
        dist_lp_itd = _gf(row, "DIST_LP_ITD")
        inc_date    = _parse_date(_gt(row, "INCEPTION_DATE")) or today
        terminal    = end_cap_cq + dist_lp_itd

        b_col = get_column_letter(2)
        c_col = get_column_letter(3)
        d_col = get_column_letter(4)
        e_col = get_column_letter(5)

        _cell(ws, r, 1, _gt(row, "INVESTOR_NAME"), align="left", fill_hex=fill)
        _cell(ws, r, 2, -contrib_itd,  fmt=_USD,    fill_hex=fill)
        _cell(ws, r, 3, terminal,      fmt=_USD,    fill_hex=fill)

        dc = ws.cell(row=r, column=4, value=inc_date)
        dc.number_format = _DATEFMT
        dc.font      = _font(size=9)
        dc.alignment = _align("left")
        dc.border    = _border()
        dc.fill      = _fill(fill)

        ec = ws.cell(row=r, column=5, value=today)
        ec.number_format = _DATEFMT
        ec.font      = _font(size=9)
        ec.alignment = _align("left")
        ec.border    = _border()
        ec.fill      = _fill(fill)

        xirr_formula = f"=XIRR({b_col}{r}:{c_col}{r},{d_col}{r}:{e_col}{r})"
        fc = ws.cell(row=r, column=6, value=xirr_formula)
        fc.number_format = "0.00%"
        fc.font      = _font(bold=True, size=9)
        fc.alignment = _align("right")
        fc.border    = _border()
        fc.fill      = _fill(_TOTAL)

        _cell(ws, r, 7, _gf(row, "GROSS_IRR") / 100.0, fmt="0.00%",
              bold=False, fill_hex=fill)

    r += 1

    # ── Section B: Net IRR ────────────────────────────────────────────────────
    r += 1
    _section_banner(ws, r, "B.  NET IRR — Post-Fee Cashflows (fees deducted from terminal value)", 10)
    r += 1

    hdrs_b = ["Investor", "CF at Inception (−)", "Net CF at Terminal (+)",
              "Inception Date", "Terminal Date",
              "Net XIRR (formula)", "PCAP Net IRR (check)"]
    for ci, (h, w) in enumerate(zip(hdrs_b, widths), 1):
        _hdr(ws, r, ci, h, width=w)

    for i, (_, row) in enumerate(pcap.iterrows()):
        r += 1
        fill = _GRAY1 if i % 2 else _GRAY2
        contrib_itd   = _gf(row, "CONTRIB_ITD")
        end_cap_cq    = _gf(row, "END_CAP_CQ")
        dist_lp_itd   = _gf(row, "DIST_LP_ITD")
        mgmt_fee_itd  = _gf(row, "MGMT_FEE_ITD_DLR")
        inc_fee_itd   = _gf(row, "INC_FEE_ITD")
        inc_date      = _parse_date(_gt(row, "INCEPTION_DATE")) or today
        net_terminal  = end_cap_cq + dist_lp_itd - mgmt_fee_itd - inc_fee_itd

        b_col = get_column_letter(2)
        c_col = get_column_letter(3)
        d_col = get_column_letter(4)
        e_col = get_column_letter(5)

        _cell(ws, r, 1, _gt(row, "INVESTOR_NAME"), align="left", fill_hex=fill)
        _cell(ws, r, 2, -contrib_itd,  fmt=_USD,   fill_hex=fill)
        _cell(ws, r, 3, net_terminal,  fmt=_USD,   fill_hex=fill)

        dc = ws.cell(row=r, column=4, value=inc_date)
        dc.number_format = _DATEFMT
        dc.font = _font(size=9); dc.alignment = _align("left")
        dc.border = _border(); dc.fill = _fill(fill)

        ec = ws.cell(row=r, column=5, value=today)
        ec.number_format = _DATEFMT
        ec.font = _font(size=9); ec.alignment = _align("left")
        ec.border = _border(); ec.fill = _fill(fill)

        xirr_formula = f"=XIRR({b_col}{r}:{c_col}{r},{d_col}{r}:{e_col}{r})"
        fc = ws.cell(row=r, column=6, value=xirr_formula)
        fc.number_format = "0.00%"
        fc.font = _font(bold=True, size=9)
        fc.alignment = _align("right")
        fc.border = _border()
        fc.fill = _fill(_TOTAL)

        _cell(ws, r, 7, _gf(row, "NET_IRR") / 100.0, fmt="0.00%", fill_hex=fill)


def _irr_from_ledger(ws, pcap: pd.DataFrame, cf_ledger: pd.DataFrame, r: int) -> None:
    """Horizontal layout: unique dates as columns, XIRR formula per investor."""
    df = cf_ledger.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    if "AMOUNT" not in df.columns or "INVESTOR_NAME" not in df.columns:
        _placeholder(ws, r, "CF Ledger missing INVESTOR_NAME or AMOUNT columns", 8)
        return

    df["AMOUNT"] = pd.to_numeric(df["AMOUNT"], errors="coerce").fillna(0.0)

    if "TRANSACTION_DATE" in df.columns:
        df["TRANSACTION_DATE"] = pd.to_datetime(df["TRANSACTION_DATE"], errors="coerce")
    else:
        _irr_from_pcap(ws, pcap, r)
        return

    today = date.today()

    # Sign convention: contributions = negative outflow, distributions/DRIP = positive inflow
    def _sign(txn_type: str) -> float:
        t = str(txn_type).strip()
        if t in ("Distribution", "DRIP", "Redemption"):
            return 1.0
        if t == "Contribution":
            return -1.0
        return 0.0

    # Collect all unique dates (transactions + today as terminal date)
    all_dates = sorted({d.date() for d in df["TRANSACTION_DATE"].dropna()}) + [today]

    # ── Section A: Gross IRR ──────────────────────────────────────────────────
    _section_banner(ws, r, "A.  GROSS IRR — Actual Cashflows (with terminal NAV added at report date)", len(all_dates) + 3)
    r += 1

    # Header row: dates
    _hdr(ws, r, 1, "Investor", width=26)
    date_col_start = 2
    for ci_off, dt in enumerate(all_dates):
        col_idx = date_col_start + ci_off
        c = ws.cell(row=r, column=col_idx, value=dt)
        c.number_format = _DATEFMT
        c.font      = _font(bold=True, color=_WHITE, size=9)
        c.fill      = _fill(_BLUE)
        c.alignment = _align("center")
        c.border    = _border()
        ws.column_dimensions[get_column_letter(col_idx)].width = 14
    irr_col = date_col_start + len(all_dates)
    _hdr(ws, r, irr_col,   "Gross XIRR",       width=18)
    _hdr(ws, r, irr_col+1, "PCAP Gross IRR",   width=18)
    hdr_row = r
    r += 1

    for i, (_, pcap_row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        inv_name = _gt(pcap_row, "INVESTOR_NAME")
        inv_cf = df[df["INVESTOR_NAME"] == inv_name].copy()

        cfs: dict[date, float] = {}
        for _, txn in inv_cf.iterrows():
            txn_date = txn["TRANSACTION_DATE"]
            if pd.isna(txn_date):
                continue
            d = txn_date.date()
            sign = _sign(str(txn.get("TYPE", "")))
            cfs[d] = cfs.get(d, 0.0) + sign * float(txn["AMOUNT"])

        # Add terminal NAV as final positive inflow at today
        cfs[today] = cfs.get(today, 0.0) + _gf(pcap_row, "END_CAP_CQ")

        _cell(ws, r, 1, inv_name, align="left", fill_hex=fill)
        cf_cells = []
        for ci_off, dt in enumerate(all_dates):
            col_idx = date_col_start + ci_off
            val = cfs.get(dt, 0.0)
            _cell(ws, r, col_idx, val, fmt=_USD, fill_hex=fill)
            cf_cells.append(get_column_letter(col_idx))

        # XIRR formula referencing this row's CF cells and the header row's date cells
        cf_range   = f"{cf_cells[0]}{r}:{cf_cells[-1]}{r}"
        date_range = f"{cf_cells[0]}{hdr_row}:{cf_cells[-1]}{hdr_row}"
        fc = ws.cell(row=r, column=irr_col, value=f"=XIRR({cf_range},{date_range})")
        fc.number_format = "0.00%"
        fc.font = _font(bold=True, size=9); fc.alignment = _align("right")
        fc.border = _border(); fc.fill = _fill(_TOTAL)

        _cell(ws, r, irr_col+1, _gf(pcap_row, "GROSS_IRR") / 100.0, fmt="0.00%", fill_hex=fill)
        r += 1

    # ── Section B: Net IRR ────────────────────────────────────────────────────
    r += 2
    _section_banner(ws, r, "B.  NET IRR — Cashflows Net of Fees (terminal NAV net of total fees)", len(all_dates) + 3)
    r += 1

    _hdr(ws, r, 1, "Investor", width=26)
    for ci_off, dt in enumerate(all_dates):
        col_idx = date_col_start + ci_off
        c = ws.cell(row=r, column=col_idx, value=dt)
        c.number_format = _DATEFMT
        c.font = _font(bold=True, color=_WHITE, size=9)
        c.fill = _fill(_BLUE)
        c.alignment = _align("center")
        c.border = _border()
    _hdr(ws, r, irr_col,   "Net XIRR",       width=18)
    _hdr(ws, r, irr_col+1, "PCAP Net IRR",   width=18)
    hdr_row2 = r
    r += 1

    for i, (_, pcap_row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        inv_name = _gt(pcap_row, "INVESTOR_NAME")
        inv_cf = df[df["INVESTOR_NAME"] == inv_name].copy()

        cfs: dict[date, float] = {}
        for _, txn in inv_cf.iterrows():
            txn_date = txn["TRANSACTION_DATE"]
            if pd.isna(txn_date):
                continue
            d = txn_date.date()
            sign = _sign(str(txn.get("TYPE", "")))
            cfs[d] = cfs.get(d, 0.0) + sign * float(txn["AMOUNT"])

        mgmt_fee = _gf(pcap_row, "MGMT_FEE_ITD_DLR")
        inc_fee  = _gf(pcap_row, "INC_FEE_ITD")
        net_nav  = _gf(pcap_row, "END_CAP_CQ") - mgmt_fee - inc_fee
        cfs[today] = cfs.get(today, 0.0) + net_nav

        _cell(ws, r, 1, inv_name, align="left", fill_hex=fill)
        cf_cells = []
        for ci_off, dt in enumerate(all_dates):
            col_idx = date_col_start + ci_off
            val = cfs.get(dt, 0.0)
            _cell(ws, r, col_idx, val, fmt=_USD, fill_hex=fill)
            cf_cells.append(get_column_letter(col_idx))

        cf_range   = f"{cf_cells[0]}{r}:{cf_cells[-1]}{r}"
        date_range = f"{cf_cells[0]}{hdr_row2}:{cf_cells[-1]}{hdr_row2}"
        fc = ws.cell(row=r, column=irr_col, value=f"=XIRR({cf_range},{date_range})")
        fc.number_format = "0.00%"
        fc.font = _font(bold=True, size=9); fc.alignment = _align("right")
        fc.border = _border(); fc.fill = _fill(_TOTAL)

        _cell(ws, r, irr_col+1, _gf(pcap_row, "NET_IRR") / 100.0, fmt="0.00%", fill_hex=fill)
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Sheet 10: Stress_IRR
# ═══════════════════════════════════════════════════════════════════════════════

def _build_stress_irr(ws, pcap: pd.DataFrame) -> None:
    ws.title = "Stress_IRR"
    r = _banner(ws, "Stress IRR — NAV Haircut Scenarios",
                "Section A: Stressed NAV | Section B: Stressed IRR | Section C: Stressed MOIC", n_cols=8)

    today = date.today()
    stress_levels = [0.0, -0.10, -0.20, -0.30, -0.15]
    stress_labels = ["Base (0%)", "−10%", "−20%", "−30%", "Custom (−15%)"]

    # ── Section A: Stressed NAV ───────────────────────────────────────────────
    _section_banner(ws, r, "A.  Stressed NAV (Ending NAV × Haircut Factor)", 7)
    r += 1

    hdrs_a   = ["Investor", "Ending NAV (Base)"] + stress_labels[1:]
    widths_a = [26, 20, 18, 18, 18, 18]
    for ci, (h, w) in enumerate(zip(hdrs_a, widths_a), 1):
        _hdr(ws, r, ci, h, width=w)
    r += 1

    nav_rows: list[list] = []
    for i, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        end_nav = _gf(row, "END_CAP_CQ")
        stressed = [end_nav * (1 + lvl) for lvl in stress_levels]
        nav_rows.append(stressed)
        _cell(ws, r, 1, _gt(row, "INVESTOR_NAME"), align="left", fill_hex=fill)
        for ci, val in enumerate(stressed, 2):
            _cell(ws, r, ci, val, fmt=_USD, bold=(ci == 2), fill_hex=fill)
        r += 1

    # ── Section B: Stressed IRR ───────────────────────────────────────────────
    r += 1
    _section_banner(ws, r, "B.  Stressed IRR  — (Stressed NAV + Dist ITD) / Contrib ITD)^(365/HoldDays) − 1", 7)
    r += 1

    hdrs_b = ["Investor"] + stress_labels
    for ci, (h, w) in enumerate(zip(hdrs_b, [26]+[18]*5), 1):
        _hdr(ws, r, ci, h, width=w)
    r += 1

    for i, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        contrib_itd = _gf(row, "CONTRIB_ITD")
        dist_lp_itd = _gf(row, "DIST_LP_ITD")
        inc_date    = _parse_date(_gt(row, "INCEPTION_DATE"))
        hold_days   = (today - inc_date).days if inc_date else 1825

        stressed_navs = nav_rows[i]
        _cell(ws, r, 1, _gt(row, "INVESTOR_NAME"), align="left", fill_hex=fill)
        for ci, (snav, lvl) in enumerate(zip(stressed_navs, stress_levels), 2):
            irr = _stress_irr(snav, dist_lp_itd, contrib_itd, hold_days)
            bold = (ci == 2)
            _cell(ws, r, ci, irr, fmt="0.00%", bold=bold, fill_hex=_TOTAL if bold else fill)
        r += 1

    # ── Section C: Stressed MOIC ──────────────────────────────────────────────
    r += 1
    _section_banner(ws, r, "C.  Stressed MOIC  — (Stressed NAV + Dist ITD) / Contrib ITD", 7)
    r += 1

    hdrs_c = ["Investor"] + stress_labels
    for ci, (h, w) in enumerate(zip(hdrs_c, [26]+[18]*5), 1):
        _hdr(ws, r, ci, h, width=w)
    r += 1

    for i, (_, row) in enumerate(pcap.iterrows()):
        fill = _GRAY1 if i % 2 else _GRAY2
        contrib_itd = _gf(row, "CONTRIB_ITD")
        dist_lp_itd = _gf(row, "DIST_LP_ITD")
        stressed_navs = nav_rows[i]
        _cell(ws, r, 1, _gt(row, "INVESTOR_NAME"), align="left", fill_hex=fill)
        for ci, snav in enumerate(stressed_navs, 2):
            moic = ((snav + dist_lp_itd) / contrib_itd) if contrib_itd > 0 else 0.0
            bold = (ci == 2)
            _cell(ws, r, ci, moic, fmt=_X, bold=bold, fill_hex=_TOTAL if bold else fill)
        r += 1

    # ── Legend ────────────────────────────────────────────────────────────────
    r += 1
    _section_banner(ws, r, "Notes & Methodology", 7)
    r += 1
    notes = [
        "Stressed IRR formula: (Stressed NAV + LP Distributions ITD) / Contrib. ITD)^(365 / Hold Days) − 1",
        "Hold Days = calendar days from Inception Date to today's date",
        "MOIC = (Stressed NAV + LP Distributions ITD) / Total Capital Contributions ITD",
        "Base (0%) = no haircut; values match PCAP Ending NAV",
        "Custom stress level set at −15%; adjust the NAV manually for other scenarios",
    ]
    for note in notes:
        ws.merge_cells(f"A{r}:H{r}")
        c = ws.cell(row=r, column=1, value=note)
        c.font = _font(italic=True, size=9)
        c.alignment = _align("left")
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Public entry point
# ═══════════════════════════════════════════════════════════════════════════════

def build_hf_workbook(pcap_df: pd.DataFrame,
                      cf_ledger_df: pd.DataFrame | None = None) -> bytes:
    """Build the 10-sheet HF PCAP Excel workbook.

    Args:
        pcap_df:      Clean investor PCAP from hf_pcap_engine.read_hf_pcap().
        cf_ledger_df: Optional transaction-level CF ledger (INVESTOR_NAME,
                      TRANSACTION_DATE, TYPE, AMOUNT, optional: QUARTER,
                      SUB_TYPE, UNITS, NOTES).

    Returns:
        .xlsx bytes ready for st.download_button.
    """
    wb = Workbook()

    sheet_names = [
        "Period_Params", "Dashboard", "Investor_Register", "PCAP",
        "Capital_Accounts", "CF_Ledger", "CF_Aggregator",
        "Distribution_Waterfall", "Cashflow_IRR", "Stress_IRR",
    ]
    wb.active.title = sheet_names[0]
    for name in sheet_names[1:]:
        wb.create_sheet(title=name)

    _build_period_params(wb["Period_Params"],          pcap_df)
    _build_dashboard(wb["Dashboard"],                  pcap_df)
    _build_investor_register(wb["Investor_Register"],  pcap_df)
    _build_pcap(wb["PCAP"],                            pcap_df)
    _build_capital_accounts(wb["Capital_Accounts"],    pcap_df)
    _build_cf_ledger(wb["CF_Ledger"],                  cf_ledger_df)
    _build_cf_aggregator(wb["CF_Aggregator"],          pcap_df, cf_ledger_df)
    _build_distribution_waterfall(wb["Distribution_Waterfall"], pcap_df)
    _build_cashflow_irr(wb["Cashflow_IRR"],            pcap_df, cf_ledger_df)
    _build_stress_irr(wb["Stress_IRR"],                pcap_df)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
