"""
KPMG Capital Statement Generator — Streamlit UI
Run: streamlit run app.py
"""

import io
import os
import zipfile
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

import pandas as pd
import streamlit as st
from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

from google import genai as _genai

_GEMINI_KEY = os.environ.get("GEMINI_API_KEY") or ""
if not _GEMINI_KEY:
    raise EnvironmentError(
        "GEMINI_API_KEY environment variable is not set. "
        "Please set it in your environment or create a .env file with GEMINI_API_KEY=your_key."
    )
_GEMINI_MODEL = "models/gemini-2.5-pro"

from generate_capital_statements import (
    build_document, build_summary_excel,
    coerce_transfer_case_nulls, _norm_label, _strip_brackets,
    fmt_usd, fmt_ratio, fmt_date,
    REQUIRED_COLS as _REQUIRED_COLS,
)

from data_wrangler import wrangle
from audit_trail import start_run, log_event, close_run
from validation_agent import validate_row

from hf_pcap_engine import read_hf_pcap_from_upload, HF_REQUIRED_COLS as _HF_REQUIRED_COLS
from hf_statement_generator import build_hf_docx, build_hf_pdf
from hf_excel_generator import build_hf_workbook

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Capital Analysis Statement Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── KPMG Brand tokens (Apex Web UI Design Guide v1.2) ────────────────────────
KPMG_BLUE  = "#00338D"   # primary brand / section headers
COBALT     = "#1E49E2"   # buttons, interactive elements
PACIFIC    = "#00B8F5"   # accent strips, active indicators
LIGHT_BLUE = "#ACEAFF"   # banner tints (light backgrounds only)
DARK_NAVY  = "#0C233C"   # sidebar, primary body text
PURPLE     = "#7213EA"   # sparingly: tags, callouts
TEAL       = "#00A3A1"   # charts (secondary series)
WHITE      = "#FFFFFF"
GRAY_1     = "#F3F6FA"   # subtle panel / row backgrounds
GRAY_2     = "#E1E6EF"   # borders, dividers
TEXT_GRAY  = "#45556B"   # secondary / muted text
SUCCESS    = "#00AB6B"   # semantic: pass / positive
WARNING    = "#FFBB1C"   # semantic: pending / warning
DANGER     = "#E63946"   # semantic: error / fail

# ── PDF styles (module-level to avoid ReportLab registry collisions) ──────────
_BASE  = getSampleStyleSheet()
_RED_C = rl_colors.HexColor("#D73B3E")

_PS_CONF     = ParagraphStyle("kpmg_conf",     parent=_BASE["Normal"], fontSize=9,  textColor=_RED_C,  alignment=TA_RIGHT, fontName="Helvetica-Bold")
_PS_TITLE    = ParagraphStyle("kpmg_title",    parent=_BASE["Normal"], fontSize=16, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=2)
_PS_SUBTITLE = ParagraphStyle("kpmg_subtitle", parent=_BASE["Normal"], fontSize=12, alignment=TA_CENTER, spaceAfter=4)
_PS_NORMAL   = ParagraphStyle("kpmg_normal",   parent=_BASE["Normal"], fontSize=11, leading=15)
_PS_INDENT   = ParagraphStyle("kpmg_indent",   parent=_BASE["Normal"], fontSize=11, leading=15, leftIndent=14)
_PS_VALUE    = ParagraphStyle("kpmg_value",    parent=_BASE["Normal"], fontSize=11, leading=15, alignment=TA_RIGHT)
_PS_FOOT     = ParagraphStyle("kpmg_foot",     parent=_BASE["Normal"], fontSize=9,  fontName="Helvetica-Oblique")


# ── PDF builder (mirrors build_document's formatting rules exactly) ──────────
_PDF_KB_BORDER = rl_colors.HexColor("#00338D")
_ZERO_EPS      = 0.005  # Change 1 threshold — matches build_document


def build_pdf_document(row: pd.Series) -> bytes:
    # Change 6: null→0 coercion for blank transfer-case columns (shared helper)
    row = coerce_transfer_case_nulls(row)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.75*inch, bottomMargin=0.75*inch,
        leftMargin=1.0*inch, rightMargin=1.0*inch,
    )
    COL_W = [4.25*inch, 2.25*inch]

    def _two_col(label: str, value: str, indented: bool = False, bordered: bool = False) -> Table:
        lbl_s = _PS_INDENT if indented else _PS_NORMAL
        data  = [[Paragraph(label, lbl_s), Paragraph(value, _PS_VALUE)]]
        t = Table(data, colWidths=COL_W)
        style_cmds = [
            ("TOPPADDING",    (0,0), (-1,-1), 1),
            ("BOTTOMPADDING", (0,0), (-1,-1), 1),
            ("LEFTPADDING",   (0,0), (-1,-1), 0),
            ("RIGHTPADDING",  (-1,0), (-1,-1), 0),
        ]
        if bordered:
            # Change 2: top+bottom KPMG-blue border on the amount (value) column only
            style_cmds += [
                ("LINEABOVE", (1, 0), (1, 0), 0.75, _PDF_KB_BORDER),
                ("LINEBELOW", (1, 0), (1, 0), 0.75, _PDF_KB_BORDER),
            ]
        t.setStyle(TableStyle(style_cmds))
        return t

    def _section(text):
        return Paragraph(f"<b><u>{text}</u></b>", _PS_NORMAL)

    def _header(label, value):
        return Paragraph(f"<b>{label}</b> {value}", _PS_NORMAL)

    sp = lambda h=6: Spacer(1, h)
    story = []

    story += [Paragraph("CONFIDENTIAL", _PS_CONF), sp(4)]
    story += [Paragraph("Capital Analysis", _PS_TITLE), sp(4)]
    story += [Paragraph(f"for the period ended {fmt_date(row['TO_DATE'])}", _PS_SUBTITLE), sp(10)]
    story += [_header("Partnership:", str(row["PARTNERSHIP_NAME"])), sp(4)]
    story += [_header("Investor:",    str(row["INVESTOR_NAME"])),    sp(4)]
    story += [_header("Investor ID:", str(row["INVESTOR_ID"])),      sp(4)]
    story += [_header("Currency:",    str(row["CURRENCY_CODE"])),    sp(12)]

    # ── Section 1: Summary of Capital Account ────────────────────────────────
    story += [_section("Summary of Capital Account"), sp(6)]

    _v = float(row["OPENING_YTD_NAV"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label(f"Opening Capital balance as on {fmt_date(row['FROM_DATE'])}"), fmt_usd(_v)))

    _v = float(row["YTD_CONTRIBUTION"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label("Capital contributions during the year"), fmt_usd(_v)))

    _v = float(row["YTD_DISTRIBUTION"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label("Distributions during the year"), fmt_usd(_v)))

    story.append(sp(6))
    story.append(Paragraph("Net investment activity:", _PS_NORMAL))

    net_income = float(row["INVESTMENT_INCOME"]) - float(row["INVESTMENT_EXPENSE"])
    if abs(net_income) >= _ZERO_EPS:
        story.append(_two_col(_norm_label("Investment and other income"), fmt_usd(net_income), indented=True))

    _v = float(row["UNREALIZED_GAINS_LOSS"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label("Net unrealized appreciation (depreciation)"), fmt_usd(_v), indented=True))

    _v = float(row["REALIZED_GAINS_LOSS"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label("Net realized gain (loss)"), fmt_usd(_v), indented=True))

    story.append(sp(6))

    _v = float(row["MANAGEMENT_FEE"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label("Management fees for the period"), fmt_usd(_v)))

    _v = float(row["INCENTIVE_FEE"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label("Incentive fees for the period"), fmt_usd(_v)))

    story.append(sp(6))

    # Change 2: double border on Ending NAV (amount cell only)
    _v = float(row["CLOSING_YTD_NAV"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(
            _norm_label(f"Capital balance (remaining value) at {fmt_date(row['TO_DATE'])} *"),
            fmt_usd(_v), bordered=True,
        ))

    story.append(sp(16))

    # ── Section 2: Summary of Capital Commitment ─────────────────────────────
    story += [_section("Summary of Capital Commitment"), sp(6)]

    committed   = float(row["COMMITTED_CAPITAL"])
    contributed = float(row["INCEPTION_TO_DATE_CONTRIBUTION"])
    remaining   = committed - contributed

    # Change 3: _strip_brackets removes ( ) from labels in this section
    if abs(committed) >= _ZERO_EPS:
        story.append(_two_col(_norm_label(_strip_brackets("Capital commitment per subscription agreement (A)")), fmt_usd(committed)))

    if abs(contributed) >= _ZERO_EPS:
        story.append(_two_col(_norm_label(_strip_brackets("Capital contributed to date (B)")), fmt_usd(contributed)))

    # Change 2: double border on Unfunded Commitment (amount cell only)
    if abs(remaining) >= _ZERO_EPS:
        story.append(_two_col(_norm_label(_strip_brackets("Remaining capital commitment (A-B)")), fmt_usd(remaining), bordered=True))

    story.append(sp(16))

    # ── Section 3: Summary of Distributions and Valuation ───────────────────
    story += [_section("Summary of Distributions and Valuation"), sp(6)]

    # Change 3: _strip_brackets removes ( ) from labels in this section
    _v = float(row["INCEPTION_TO_DATE_CONTRIBUTION"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label(_strip_brackets("Total capital contributed to date")), fmt_usd(_v)))

    _v = float(row["INCEPTION_TO_DATE_DISTRIBUTION"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label(_strip_brackets("Total distributions to date")), fmt_usd(_v)))

    _v = float(row["CLOSING_YTD_NAV"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(
            _norm_label(_strip_brackets(f"Capital balance (remaining value) at {fmt_date(row['TO_DATE'])} *")),
            fmt_usd(_v),
        ))

    # Change 2: double border on TEV (amount cell only)
    _v = float(row["TEV"])
    if abs(_v) >= _ZERO_EPS:
        story.append(_two_col(_norm_label(_strip_brackets("Total Estimated Value (distributions + balance)")), fmt_usd(_v), bordered=True))

    # Change 2: double border on TEV Ratio — no zero-suppression (ratio, not monetary)
    story.append(_two_col(_norm_label(_strip_brackets("Total Estimated Value as net multiple")), fmt_ratio(row["TEV_RATIO"]), bordered=True))

    story.append(sp(20))
    story.append(Paragraph(
        "* Represents remaining value. The remaining value is based upon available "
        "information and may not represent amounts which might ultimately be realized.",
        _PS_FOOT,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── AI: system prompt + insight types ─────────────────────────────────────────
_PE_SYSTEM = """\
You are a senior Private Equity investment advisor with 20+ years of experience at top-tier global PE firms \
(Blackstone, KKR, Apollo, Carlyle). You specialise in analyzing LP capital accounts and fund performance to \
provide strategic insights for HNIs, family offices, and institutional investors.

Your expertise covers LP/GP dynamics, DPI/RVPI/TVPI, fee analysis, distribution planning, capital deployment \
pacing, exit strategy, and risk-adjusted return benchmarking (top-quartile PE: TVPI >2.0x; median ~1.6x).

Be direct, data-driven, and structured. Use PE terminology. Flag red flags clearly. \
Every insight must be tied to the specific numbers provided — no generic commentary.\
"""

_INSIGHT_INSTRUCTIONS = {
    "Full Investment Analysis": (
        "Cover: (1) Performance vs PE benchmarks (TVPI/DPI/RVPI), "
        "(2) Capital account health, (3) Fee drag on net returns, "
        "(4) YTD activity, (5) Key risks and opportunities, (6) Specific recommendations."
    ),
    "Fee & Cost Analysis": (
        "Cover: (1) Management fee as % of committed/contributed vs industry norm (1.5–2%), "
        "(2) Incentive fee vs value created, (3) Total fee drag on gross-to-net return, "
        "(4) Fee-adjusted DPI and TVPI, (5) Recommendations."
    ),
    "Distribution & Liquidity Planning": (
        "Cover: (1) DPI and capital return timeline, (2) RVPI realization outlook, "
        "(3) YTD distribution pace, (4) Liquidity needs and secondary market optionality, "
        "(5) Distribution strategy recommendation."
    ),
    "Exit Strategy Assessment": (
        "Cover: (1) TVPI vs typical PE exit multiples, (2) Unrealized gain and exit readiness, "
        "(3) NAV trajectory signals, (4) Optimal hold period, "
        "(5) Exit scenario modeling (base/bull/bear), (6) Timing recommendation."
    ),
    "Commitment Utilization & Pacing": (
        "Cover: (1) Utilization rate vs typical PE pace (60–80% by year 3–4), "
        "(2) Unfunded commitment and remaining call risk, "
        "(3) Over-commitment risk, (4) Capital call timing projections, "
        "(5) Recommendations for managing remaining obligations."
    ),
}


_HF_SYSTEM = """\
You are a senior Hedge Fund analyst with 20+ years of experience at top-tier hedge funds \
(Citadel, Bridgewater, DE Shaw, Millennium). You specialise in hedge fund LP capital accounts, \
PCAP analysis, unit-based NAV attribution, and risk-adjusted returns.

Your expertise covers unit pricing, NAV attribution, IRR vs. TWR, management/incentive fee drag, \
waterfall mechanics, stress testing, and LP redemption risk.

Be direct, data-driven, and structured. Use hedge fund terminology precisely. \
Every insight must be tied to the specific numbers provided — no generic commentary.\
"""

_HF_INSIGHT_INSTRUCTIONS = {
    "Full HF Performance Analysis": (
        "Cover: (1) NAV and unit price movement attribution, (2) IRR vs. hurdle, "
        "(3) Fee drag (management + incentive), (4) Capital account composition, "
        "(5) Stress test implications, (6) Key risks and strategic recommendations."
    ),
    "Fee & Carry Analysis": (
        "Cover: (1) Incentive fee as % of gross return, (2) Hurdle rate adequacy, "
        "(3) GP carry vs. LP net return split, (4) Fee-adjusted IRR, "
        "(5) Comparison to industry norms (2-and-20 vs actual structure), (6) Recommendations."
    ),
    "Waterfall & Distribution Analysis": (
        "Cover: (1) Excess return above hurdle, (2) GP vs LP distribution split, "
        "(3) Effective carry rate, (4) LP net return after carry, "
        "(5) Comparison to contributed capital, (6) Pacing of distributions."
    ),
    "Stress Test Assessment": (
        "Cover: (1) NAV sensitivity at -10%/-20%/-30% haircuts, (2) IRR impact under each scenario, "
        "(3) Break-even analysis, (4) Probability-weighted return, "
        "(5) Liquidity and redemption risk, (6) Hedging / risk mitigation recommendations."
    ),
    "Unit Price & NAV Attribution": (
        "Cover: (1) Unit price change decomposition (income / unrealized / realized / expenses / fees), "
        "(2) Investor share vs. fund-level P&L, (3) DRIP impact on units and NAV, "
        "(4) Redemption dilution effect, (5) Ending NAV quality assessment."
    ),
}


def _safe_div(num, den):
    try:
        return float(num) / float(den) if float(den) != 0 else 0.0
    except (TypeError, ValueError):
        return 0.0


def _portfolio_context(df_sel: pd.DataFrame, partnership: str) -> str:
    total_committed   = df_sel["COMMITTED_CAPITAL"].sum()
    total_contributed = df_sel["INCEPTION_TO_DATE_CONTRIBUTION"].sum()
    total_dist        = df_sel["INCEPTION_TO_DATE_DISTRIBUTION"].sum()
    total_nav         = df_sel["CLOSING_YTD_NAV"].sum()
    total_tev         = df_sel["TEV"].sum()
    tvpi  = _safe_div(total_tev, total_contributed)
    dpi   = _safe_div(total_dist, total_contributed)
    rvpi  = _safe_div(total_nav, total_contributed)
    util  = _safe_div(total_contributed, total_committed) * 100

    lines = [
        f"Partnership: {partnership}",
        f"LPs: {len(df_sel)}",
        f"Total Committed: {fmt_usd(total_committed)}",
        f"Total Contributed ITD: {fmt_usd(total_contributed)}",
        f"Total Distributions ITD: {fmt_usd(total_dist)}",
        f"Portfolio NAV: {fmt_usd(total_nav)}",
        f"Portfolio TEV: {fmt_usd(total_tev)}",
        f"TVPI: {tvpi:.2f}x | DPI: {dpi:.2f}x | RVPI: {rvpi:.2f}x",
        f"Commitment Utilization: {util:.1f}%",
        "LP positions:",
    ]
    for _, r in df_sel.iterrows():
        lines.append(
            f"  {r['INVESTOR_NAME']}: Committed {fmt_usd(r['COMMITTED_CAPITAL'])}, "
            f"NAV {fmt_usd(r['CLOSING_YTD_NAV'])}, TVPI {fmt_ratio(r['TEV_RATIO'])}"
        )
    return "\n".join(lines)


def _build_portfolio_prompt(df_sel, partnership, as_of, insight_type):
    ctx = _portfolio_context(df_sel, partnership)
    instr = _INSIGHT_INSTRUCTIONS.get(insight_type, _INSIGHT_INSTRUCTIONS["Full Investment Analysis"])
    return f"AS OF: {as_of}\n\n{ctx}\n\nANALYSIS REQUEST (Portfolio Level):\n{instr}"


def _build_investor_prompt(row: pd.Series, insight_type: str) -> str:
    committed   = float(row["COMMITTED_CAPITAL"])
    contributed = float(row["INCEPTION_TO_DATE_CONTRIBUTION"])
    net_income  = float(row["INVESTMENT_INCOME"]) - float(row["INVESTMENT_EXPENSE"])
    dpi  = _safe_div(row["INCEPTION_TO_DATE_DISTRIBUTION"], contributed)
    rvpi = _safe_div(row["CLOSING_YTD_NAV"], contributed)

    ctx = (
        f"Investor: {row['INVESTOR_NAME']} | Partnership: {row['PARTNERSHIP_NAME']}\n"
        f"Period: {fmt_date(row['FROM_DATE'])} – {fmt_date(row['TO_DATE'])} | CCY: {row['CURRENCY_CODE']}\n"
        f"Committed: {fmt_usd(committed)} | Contributed ITD: {fmt_usd(contributed)} | "
        f"Remaining: {fmt_usd(committed - contributed)} | Utilization: {_safe_div(contributed, committed)*100:.1f}%\n"
        f"Opening NAV: {fmt_usd(row['OPENING_YTD_NAV'])} | Closing NAV: {fmt_usd(row['CLOSING_YTD_NAV'])}\n"
        f"YTD Contributions: {fmt_usd(row['YTD_CONTRIBUTION'])} | YTD Distributions: {fmt_usd(row['YTD_DISTRIBUTION'])}\n"
        f"Net Income: {fmt_usd(net_income)} | Unrealized G/L: {fmt_usd(row['UNREALIZED_GAINS_LOSS'])} | "
        f"Realized G/L: {fmt_usd(row['REALIZED_GAINS_LOSS'])}\n"
        f"Mgmt Fee: {fmt_usd(row['MANAGEMENT_FEE'])} | Incentive Fee: {fmt_usd(row['INCENTIVE_FEE'])}\n"
        f"TEV: {fmt_usd(row['TEV'])} | TVPI: {fmt_ratio(row['TEV_RATIO'])} | DPI: {dpi:.2f}x | RVPI: {rvpi:.2f}x\n"
        f"Distributions ITD: {fmt_usd(row['INCEPTION_TO_DATE_DISTRIBUTION'])}"
    )
    instr = _INSIGHT_INSTRUCTIONS.get(insight_type, _INSIGHT_INSTRUCTIONS["Full Investment Analysis"])
    return f"{ctx}\n\nANALYSIS REQUEST (Investor Level):\n{instr}"


def _build_hf_portfolio_prompt(pcap_df: pd.DataFrame, insight_type: str) -> str:
    def _col(col, fn):
        if col in pcap_df.columns:
            vals = pcap_df[col].dropna()
            return fn(vals) if len(vals) else 0.0
        return 0.0

    total_nav   = _col("END_CAP_CQ",  sum)
    total_itd   = _col("END_CAP_ITD", sum)
    total_cont  = _col("CONTRIB_ITD", sum)
    total_dist  = _col("DIST_LP_ITD", sum)
    avg_irr_g   = _col("GROSS_IRR",   lambda x: x.mean())
    avg_irr_n   = _col("NET_IRR",     lambda x: x.mean())
    avg_tvpi    = _col("TVPI",        lambda x: x.mean())
    avg_dpi     = _col("DPI",         lambda x: x.mean())
    avg_rvpi    = _col("RVPI",        lambda x: x.mean())
    tot_inc_fee = _col("INC_FEE_ITD", sum)
    tot_lp_net  = _col("LP_NET_WF",   sum)
    beg_px      = _col("BEG_PX",      lambda x: x.mean())
    end_px      = _col("END_PX",      lambda x: x.mean())

    lines = [
        f"Investors (LPs): {len(pcap_df)}",
        f"Avg Beginning Unit Price: {fmt_usd(beg_px)}  |  Avg Ending Unit Price: {fmt_usd(end_px)}",
        f"Total NAV (CQ Ending): {fmt_usd(total_nav)}",
        f"Total NAV (ITD Ending): {fmt_usd(total_itd)}",
        f"Total Contributions (ITD): {fmt_usd(total_cont)}",
        f"Total Distributions to LPs (ITD): {fmt_usd(total_dist)}",
        f"Total Incentive Fees (ITD): {fmt_usd(tot_inc_fee)}",
        f"Total LP Net Waterfall Share: {fmt_usd(tot_lp_net)}",
        f"Avg Gross IRR: {avg_irr_g:.2f}%  |  Avg Net IRR: {avg_irr_n:.2f}%",
        f"Avg TVPI: {avg_tvpi:.2f}x  |  Avg DPI: {avg_dpi:.2f}x  |  Avg RVPI: {avg_rvpi:.2f}x",
        "LP positions:",
    ]
    for _, r in pcap_df.iterrows():
        nav_cq = float(r.get("END_CAP_CQ", 0) or 0)
        irr_n  = float(r.get("NET_IRR", 0) or 0)
        tvpi   = float(r.get("TVPI", 0) or 0)
        lp_net = float(r.get("LP_NET_WF", 0) or 0)
        lines.append(
            f"  {r.get('INVESTOR_NAME','?')}: NAV(CQ) {fmt_usd(nav_cq)}, "
            f"Net IRR {irr_n:.2f}%, TVPI {tvpi:.2f}x, LP Net WF {fmt_usd(lp_net)}"
        )

    ctx   = "\n".join(lines)
    instr = _HF_INSIGHT_INSTRUCTIONS.get(insight_type, _HF_INSIGHT_INSTRUCTIONS["Full HF Performance Analysis"])
    return f"{ctx}\n\nANALYSIS REQUEST (Portfolio Level):\n{instr}"


def _build_hf_investor_prompt(row: pd.Series, insight_type: str) -> str:
    def _f(col, dft=0.0):
        v = row.get(col, dft)
        try:
            return float(v) if v is not None and not (isinstance(v, float) and pd.isna(v)) else dft
        except (TypeError, ValueError):
            return dft

    def _s(col, dft="—"):
        v = row.get(col, dft)
        s = str(v).strip() if v is not None else ""
        return s if s and s not in ("nan", "None", "") else dft

    ctx = (
        f"Investor: {_s('INVESTOR_NAME')}  |  Inception: {_s('INCEPTION_DATE')}  |  Currency: {_s('REPT_CCY')}\n"
        f"Unit Prices: Beginning {fmt_usd(_f('BEG_PX'))} → Ending {fmt_usd(_f('END_PX'))}\n"
        f"Capital (CQ): {fmt_usd(_f('BEG_CAP_CQ'))} → {fmt_usd(_f('END_CAP_CQ'))}\n"
        f"Capital (YTD): {fmt_usd(_f('BEG_CAP_YTD'))} → {fmt_usd(_f('END_CAP_YTD'))}\n"
        f"Capital (ITD): {fmt_usd(_f('BEG_CAP_ITD'))} → {fmt_usd(_f('END_CAP_ITD'))}\n"
        f"Contributions ITD: {fmt_usd(_f('CONTRIB_ITD'))}  |  Distributions to LP ITD: {fmt_usd(_f('DIST_LP_ITD'))}\n"
        f"Investment Income (CQ/YTD/ITD): {fmt_usd(_f('INC_CQ'))} / {fmt_usd(_f('INC_YTD'))} / {fmt_usd(_f('INC_ITD'))}\n"
        f"Unrealized G/L (CQ/ITD): {fmt_usd(_f('UNRLZ_CQ'))} / {fmt_usd(_f('UNRLZ_ITD'))}\n"
        f"Realized G/L (CQ/ITD): {fmt_usd(_f('RLZD_CQ'))} / {fmt_usd(_f('RLZD_ITD'))}\n"
        f"Incentive Fee (CQ/YTD/ITD): {fmt_usd(_f('INC_FEE_CQ'))} / {fmt_usd(_f('INC_FEE_YTD'))} / {fmt_usd(_f('INC_FEE_ITD'))}\n"
        f"Gross IRR: {_f('GROSS_IRR'):.2f}%  |  Net IRR: {_f('NET_IRR'):.2f}%\n"
        f"DPI: {_f('DPI'):.2f}x  |  RVPI: {_f('RVPI'):.2f}x  |  TVPI: {_f('TVPI'):.2f}x\n"
        f"Total Return (CQ/ITD): {fmt_usd(_f('TOT_RET_CQ_DLR'))} / {fmt_usd(_f('TOT_RET_ITD_DLR'))}\n"
        f"Net Return ITD: {fmt_usd(_f('NET_RET_ITD_DLR'))}\n"
        f"Hurdle Amt (ITD): {fmt_usd(_f('HURDLE_AMT_ITD'))}  |  Excess Over Hurdle: {fmt_usd(_f('EXCESS_HURDLE'))}\n"
        f"GP Catch-Up: {fmt_usd(_f('GP_CATCHUP_AMT'))}  |  LP Net WF Share: {fmt_usd(_f('LP_NET_WF'))}\n"
        f"Total Commit: {fmt_usd(_f('TOTAL_COMMIT'))}  |  Funded: {fmt_usd(_f('FUNDED_COMMIT'))}  |  Available: {fmt_usd(_f('AVAIL_COMMIT'))}\n"
        f"Lock-up: {_f('LOCKUP_MO'):.0f} months  |  Expired: {_s('LOCKUP_EXPIRED')}  |  HWM Active: {_s('HWM_ACTIVE')}\n"
        f"Pref Return: {_f('PREF_RET'):.2f}%  |  Inc Fee Rate: {_f('INC_FEE_RATE'):.2f}%  |  Hurdle Type: {_s('HURDLE_TYPE')}"
    )
    instr = _HF_INSIGHT_INSTRUCTIONS.get(insight_type, _HF_INSIGHT_INSTRUCTIONS["Full HF Performance Analysis"])
    return f"{ctx}\n\nANALYSIS REQUEST (Investor Level):\n{instr}"


def _call_gemini(prompt: str) -> str:
    client = _genai.Client(api_key=_GEMINI_KEY)
    response = client.models.generate_content(
        model=_GEMINI_MODEL,
        config=_genai.types.GenerateContentConfig(
            system_instruction=_PE_SYSTEM,
            max_output_tokens=65536,
        ),
        contents=prompt,
    )
    return response.text or ""


def _call_gemini_chat(messages: list[dict], portfolio_ctx: str = "") -> str:
    client = _genai.Client(api_key=_GEMINI_KEY)
    system = _PE_SYSTEM
    if portfolio_ctx:
        system += f"\n\nCURRENT PORTFOLIO CONTEXT (use this when answering questions):\n{portfolio_ctx}"

    contents = []
    for m in messages:
        contents.append(
            _genai.types.Content(
                role="user" if m["role"] == "user" else "model",
                parts=[_genai.types.Part(text=m["content"])],
            )
        )
    response = client.models.generate_content(
        model=_GEMINI_MODEL,
        config=_genai.types.GenerateContentConfig(
            system_instruction=system,
            max_output_tokens=65536,
        ),
        contents=contents,
    )
    return response.text or ""


# ── Single consolidated CSS block ─────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;600;700;800&display=swap');

  /* === Base typography (Open Sans, Apex Design Guide v1.2) === */
  html, body, [class*="css"] {
    font-family: 'Open Sans', Arial, sans-serif;
    color: #0C233C;
  }
  .stApp { background: #FFFFFF; }

  [data-testid="stAppViewContainer"] > .main > .block-container {
    padding-top: 0 !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 100% !important;
  }

  /* === KPMG header bar === */
  .kpmg-header {
    background: #00338D;
    margin: 0 -2rem;
    padding: 14px 32px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .kpmg-wordmark {
    font-family: 'Open Sans', 'Arial Black', sans-serif;
    font-size: 22px;
    color: #FFFFFF;
    font-weight: 800;
    letter-spacing: 0.06em;
  }
  .kpmg-app-title {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 15px;
    color: #FFFFFF;
  }
  .kpmg-accent-strip {
    height: 6px;
    background: #00B8F5;
    margin: 0 -2rem 24px -2rem;
  }

  /* === Sidebar === */
  section[data-testid="stSidebar"] {
    background: #0C233C !important;
  }
  section[data-testid="stSidebar"] > div {
    background: #0C233C !important;
  }
  section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
  section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] span,
  section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div {
    color: rgba(255,255,255,0.80) !important;
    font-family: 'Open Sans', Arial, sans-serif !important;
  }
  section[data-testid="stSidebar"] label,
  section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {
    color: rgba(255,255,255,0.80) !important;
    font-family: 'Open Sans', Arial, sans-serif !important;
  }
  section[data-testid="stSidebar"] h1,
  section[data-testid="stSidebar"] h2,
  section[data-testid="stSidebar"] h3 {
    color: #FFFFFF !important;
    font-family: 'Open Sans', Arial, sans-serif !important;
    font-weight: 700 !important;
  }
  section[data-testid="stSidebar"] .stFileUploader {
    background: rgba(172, 234, 255, 0.08);
    border: 1px dashed rgba(172, 234, 255, 0.40);
    border-radius: 4px;
    padding: 8px;
  }
  section[data-testid="stSidebar"] .stFileUploader small,
  section[data-testid="stSidebar"] .stFileUploader span,
  section[data-testid="stSidebar"] .stFileUploader p {
    color: rgba(255,255,255,0.80) !important;
  }
  section[data-testid="stSidebar"] .stTextArea textarea {
    background: rgba(255, 255, 255, 0.10) !important;
    border: 1px solid rgba(172, 234, 255, 0.30) !important;
    color: #FFFFFF !important;
    border-radius: 4px;
    font-size: 13px !important;
    resize: none;
  }
  section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
    background: rgba(255,255,255,0.10);
    color: rgba(255,255,255,0.80);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 8px;
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 14px;
    width: 100%;
    box-shadow: none;
  }
  section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover {
    background: #00B8F5;
    color: #0C233C;
    border-color: #00B8F5;
  }
  section[data-testid="stSidebar"] .stRadio label span {
    color: rgba(255,255,255,0.80) !important;
  }
  section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] {
    background: rgba(255, 255, 255, 0.10) !important;
    border-color: rgba(172, 234, 255, 0.30) !important;
  }

  .sidebar-divider {
    border: none;
    border-top: 1px solid rgba(255,255,255,0.15);
    margin: 14px 0;
  }
  .sidebar-section-label {
    font-family: 'Open Sans', Arial, sans-serif !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.10em;
    color: #FFFFFF !important;
    margin-bottom: 6px;
  }

  /* === Chat bubbles === */
  .chat-bubble-user {
    background: rgba(255,255,255,0.18);
    border-radius: 8px;
    padding: 7px 11px;
    margin: 4px 0 4px 20px;
    font-size: 12px;
    color: #FFFFFF;
    word-wrap: break-word;
    font-family: 'Open Sans', Arial, sans-serif;
  }
  .chat-bubble-ai {
    background: rgba(255,255,255,0.08);
    border-radius: 8px;
    padding: 7px 11px;
    margin: 4px 20px 4px 0;
    font-size: 12px;
    color: rgba(255,255,255,0.80);
    word-wrap: break-word;
    font-family: 'Open Sans', Arial, sans-serif;
  }
  .chat-label-user {
    font-size: 10px;
    color: rgba(255,255,255,0.60);
    text-align: right;
    margin-bottom: 1px;
  }
  .chat-label-ai {
    font-size: 10px;
    color: rgba(255,255,255,0.60);
    margin-bottom: 1px;
  }

  /* === Metric cards (HTML grid) === */
  .metric-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin-bottom: 24px;
  }
  .metric-card {
    background: #FFFFFF;
    border: 1px solid #E1E6EF;
    border-top: 4px solid #00B8F5;
    border-radius: 6px;
    padding: 18px 20px;
    box-shadow: 0px 2px 4px rgba(0,0,0,0.05);
  }
  .metric-card .label {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #45556B;
    margin-bottom: 4px;
  }
  .metric-card .value {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 24px;
    font-weight: 700;
    color: #0C233C;
    line-height: 1.1;
  }
  .metric-card .sub {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 12px;
    color: #45556B;
    margin-top: 4px;
  }

  /* === Section headers === */
  .section-header {
    background: #00338D;
    color: #FFFFFF;
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 13px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.10em;
    padding: 8px 16px;
    margin: 24px 0 12px 0;
    border-left: 4px solid #00B8F5;
    border-radius: 0 4px 4px 0;
  }

  /* === Content cards === */
  .content-card {
    background: #F3F6FA;
    border-left: 4px solid #00B8F5;
    border-radius: 0 6px 6px 0;
    padding: 16px 20px;
    margin-bottom: 16px;
    color: #0C233C;
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 13px;
    line-height: 1.6;
  }

  /* === Investor pills === */
  .investor-pill {
    display: inline-block;
    background: #F3F6FA;
    border: 1px solid #E1E6EF;
    color: #0C233C;
    border-radius: 20px;
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 12px;
    font-weight: 600;
    padding: 3px 12px;
    margin: 3px;
  }

  /* === Main-area buttons (Cobalt Blue, Apex spec: radius 8px, weight 600) === */
  div[data-testid="stButton"] > button {
    background: #1E49E2;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    font-family: 'Open Sans', Arial, sans-serif;
    font-weight: 600;
    font-size: 14px;
    padding: 10px 20px;
    box-shadow: 0px 3px 6px rgba(0,0,0,0.10);
    width: 100%;
    transition: background 0.15s ease;
  }
  div[data-testid="stButton"] > button:hover {
    background: #00338D;
    color: #FFFFFF;
    box-shadow: 0px 4px 8px rgba(0,0,0,0.15);
  }
  div[data-testid="stButton"] > button:active {
    background: #022569;
    box-shadow: none;
  }
  div[data-testid="stButton"] > button:disabled {
    background: #B5B9C2;
    color: #FFFFFF;
    opacity: 1;
    box-shadow: none;
  }

  /* === Download buttons === */
  div[data-testid="stDownloadButton"] > button {
    background: #00B8F5;
    color: #0C233C;
    border: none;
    border-radius: 8px;
    font-family: 'Open Sans', Arial, sans-serif;
    font-weight: 600;
    font-size: 13px;
    box-shadow: 0px 3px 6px rgba(0,0,0,0.10);
    width: 100%;
    transition: background 0.15s ease;
  }
  div[data-testid="stDownloadButton"] > button:hover {
    background: #00338D;
    color: #FFFFFF;
  }

  /* === Download panel === */
  .download-panel {
    background: #F3F6FA;
    padding: 20px 24px;
    margin-bottom: 16px;
    border: 1px solid #E1E6EF;
    border-left: 4px solid #00B8F5;
    border-radius: 0 6px 6px 0;
  }
  .download-panel-title {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 13px;
    font-weight: 700;
    color: #0C233C;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-bottom: 12px;
  }

  /* === Result rows === */
  .result-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 14px;
    margin-bottom: 6px;
    font-size: 13px;
    font-family: 'Open Sans', Arial, sans-serif;
    color: #0C233C;
    background: #FFFFFF;
    border: 1px solid #E1E6EF;
    border-radius: 0 6px 6px 0;
  }
  .result-ok  { border-left: 4px solid #00AB6B; }
  .result-err { border-left: 4px solid #E63946; }

  /* === Validation badges === */
  .badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 4px;
    font-size: 11px;
    font-family: 'Open Sans', Arial, sans-serif;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  .badge-pass    { background: #00AB6B; color: #FFFFFF; }
  .badge-revalue { background: #FFBB1C; color: #0C233C; }
  .badge-fail    { background: #E63946; color: #FFFFFF; }

  /* === AI insights header === */
  .ai-header {
    background: #00338D;
    color: #FFFFFF;
    padding: 14px 20px;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 12px;
    border-left: 4px solid #00B8F5;
    border-radius: 0 6px 6px 0;
  }
  .ai-title {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 15px;
    font-weight: 700;
    color: #FFFFFF;
  }
  .ai-sub {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 11px;
    color: rgba(255,255,255,0.75);
    margin-top: 2px;
  }
  .ai-badge {
    background: #00B8F5;
    color: #0C233C;
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 9px;
    font-weight: 700;
    letter-spacing: 0.10em;
    padding: 3px 8px;
    border-radius: 4px;
    text-transform: uppercase;
    white-space: nowrap;
  }

  /* === Hedge Fund info banner === */
  .hf-banner {
    background: #F3F6FA;
    border: 1px solid #E1E6EF;
    border-left: 4px solid #00B8F5;
    border-radius: 0 6px 6px 0;
    color: #0C233C;
    padding: 14px 20px;
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 13px;
    line-height: 1.6;
    margin-bottom: 16px;
  }

  /* === Tabs === */
  [data-testid="stTabs"] button[data-testid="stTab"] {
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 14px;
    font-weight: 600;
    color: #45556B;
    padding: 8px 24px;
  }
  [data-testid="stTabs"] button[data-testid="stTab"][aria-selected="true"] {
    color: #00338D;
    font-weight: 700;
    border-bottom: 3px solid #00B8F5;
  }

  /* === Widget labels (main area) === */
  [data-testid="stWidgetLabel"] p {
    color: #0C233C !important;
    font-family: 'Open Sans', Arial, sans-serif !important;
    font-weight: 600 !important;
    font-size: 13px !important;
  }
  .stRadio label span  { color: #0C233C !important; }
  .stSelectbox label   { color: #0C233C !important; }
  .stMultiSelect label { color: #0C233C !important; }

  /* === Multiselect tags (main area) === */
  span[data-baseweb="tag"] {
    background: #F3F6FA !important;
    border: 1px solid #E1E6EF !important;
    color: #0C233C !important;
    border-radius: 4px !important;
  }

  /* === Expander === */
  [data-testid="stExpander"] {
    border: 1px solid #E1E6EF;
    border-radius: 6px;
  }
  [data-testid="stExpander"] summary {
    color: #0C233C;
    font-family: 'Open Sans', Arial, sans-serif;
    font-size: 13px;
    font-weight: 600;
  }

  /* === Progress bar === */
  .stProgress > div > div > div > div { background: #1E49E2; }

  /* === Caption / info text === */
  .stCaption { color: #45556B !important; }
  .stAlert   { color: #0C233C !important; }

  /* === Hide Streamlit chrome === */
  #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


REQUIRED_COLS = _REQUIRED_COLS

# ── Session state defaults ────────────────────────────────────────────────────
for _k, _v in {
    "gen":               None,
    "chat_history":      [],
    "show_chat":         False,
    "chat_input_key":    0,
    "pe_insights":       None,
    "pe_insights_label": "",
    "hf_gen":            None,
    "hf_insights":       None,
    "hf_insights_label": "",
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ── Sidebar block 1: KPMG branding + file upload ──────────────────────────────
with st.sidebar:
    logo_path = os.path.join(os.path.dirname(__file__), "kpmg_logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=140)
    else:
        st.markdown(
            "<div style='font-family:\"KPMG Bold\",\"Arial Black\",sans-serif;"
            "font-size:28px;font-weight:900;color:#FFFFFF;letter-spacing:0.08em;"
            "padding:8px 0 2px 0;'>KPMG</div>"
            "<div style='font-family:\"Univers Light\",Arial,sans-serif;font-size:12px;"
            "color:#ACEAFF;line-height:1.5;padding-bottom:10px;'>"
            "Capital Analysis<br>Statement Generator</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<hr class='sidebar-divider'>", unsafe_allow_html=True)
    st.markdown("<p class='sidebar-section-label'>File Upload</p>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Upload investor data",
        type=["xlsx"],
        help="Must contain all required columns.",
        label_visibility="collapsed",
    )


# ── File processing (top-level, before tabs) ──────────────────────────────────
df_raw        = None
df_latest     = None
all_investors = []
n_investors   = 0
n_periods     = 0
partnership   = "—"
chosen        = []
read_error    = None
missing: set  = set()

if uploaded_file is not None:
    try:
        df_raw  = pd.read_excel(uploaded_file)
        missing = REQUIRED_COLS - set(df_raw.columns)
        if not missing:
            df_raw["TO_DATE"]   = pd.to_datetime(df_raw["TO_DATE"],   errors="coerce")
            df_raw["FROM_DATE"] = pd.to_datetime(df_raw["FROM_DATE"], errors="coerce")
            df_latest = (
                df_raw.sort_values("TO_DATE", ascending=True)
                .groupby("INVESTOR_NAME", as_index=False)
                .last()
            )
            all_investors = sorted(df_latest["INVESTOR_NAME"].tolist())
            n_investors   = len(all_investors)
            n_periods     = df_raw["TO_DATE"].nunique()
            partnership   = df_latest["PARTNERSHIP_NAME"].iloc[0] if n_investors else "—"
    except Exception as exc:
        read_error = str(exc)


# ── Sidebar block 2: investor selection (only when file loaded) ───────────────
if df_latest is not None:
    with st.sidebar:
        st.markdown("<hr class='sidebar-divider'>", unsafe_allow_html=True)
        st.markdown("<p class='sidebar-section-label'>Investor Selection</p>", unsafe_allow_html=True)
        selection_mode = st.radio(
            "Generate for:",
            ["All investors", "Selected investors"],
            horizontal=True,
            key="sel_mode",
            label_visibility="collapsed",
        )
        if selection_mode == "Selected investors":
            chosen = st.multiselect(
                "Investors",
                options=all_investors,
                default=all_investors,
                key="sel_investors",
                label_visibility="collapsed",
            )
        else:
            chosen = all_investors


# ── Sidebar block 3: PE Chat toggle + footer ──────────────────────────────────
with st.sidebar:
    st.markdown("<hr class='sidebar-divider'>", unsafe_allow_html=True)
    st.markdown("<p class='sidebar-section-label'>PE Chat</p>", unsafe_allow_html=True)

    chat_btn_label = "Close Chat" if st.session_state["show_chat"] else "Open PE Chat"
    if st.button(chat_btn_label, key="chat_toggle"):
        st.session_state["show_chat"] = not st.session_state["show_chat"]
        st.rerun()

    if st.session_state["show_chat"]:
        st.markdown(
            "<p style='font-size:10px;color:#ACEAFF;margin:4px 0 8px;'>"
            "Ask anything about PE, your portfolio, or specific investors.</p>",
            unsafe_allow_html=True,
        )

        history = st.session_state["chat_history"]
        if history:
            msgs_html = ""
            for m in history:
                if m["role"] == "user":
                    msgs_html += (
                        f"<div class='chat-label-user'>You</div>"
                        f"<div class='chat-bubble-user'>{m['content']}</div>"
                    )
                else:
                    msgs_html += (
                        f"<div class='chat-label-ai'>Gemini PE</div>"
                        f"<div class='chat-bubble-ai'>{m['content']}</div>"
                    )
            st.markdown(msgs_html, unsafe_allow_html=True)
        else:
            st.markdown(
                "<p style='font-size:11px;color:#ACEAFF;'>No messages yet. Ask your first question below.</p>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        user_input = st.text_area(
            "Message",
            key=f"chat_msg_{st.session_state['chat_input_key']}",
            placeholder="e.g. What is a good DPI for a PE fund in year 5?",
            label_visibility="collapsed",
            height=80,
        )

        col_send, col_clear = st.columns([3, 1])
        with col_send:
            send_clicked = st.button("Send", key="chat_send_btn", use_container_width=True)
        with col_clear:
            if st.button("X", key="chat_clear_btn", use_container_width=True):
                st.session_state["chat_history"] = []
                st.rerun()

        if send_clicked and user_input.strip():
            st.session_state["chat_history"].append({"role": "user", "content": user_input.strip()})
            with st.spinner(""):
                try:
                    ctx = ""
                    if st.session_state["gen"]:
                        g   = st.session_state["gen"]
                        ctx = _portfolio_context(g["df_selected"], g["partnership"])
                    reply = _call_gemini_chat(st.session_state["chat_history"], ctx)
                    st.session_state["chat_history"].append({"role": "model", "content": reply})
                except Exception as e:
                    st.session_state["chat_history"].append(
                        {"role": "model", "content": f"Error: {e}"}
                    )
            st.session_state["chat_input_key"] += 1
            st.rerun()

    st.markdown("<hr class='sidebar-divider'>", unsafe_allow_html=True)
    st.markdown(
        "<p style='font-size:10px;color:#ACEAFF;line-height:1.6;'>"
        "Documents generated are CONFIDENTIAL.<br>For internal use only.<br><br>"
        "© KPMG International</p>",
        unsafe_allow_html=True,
    )


# ── Main area: header + accent strip ─────────────────────────────────────────
st.markdown("""
<div class="kpmg-header">
  <div class="kpmg-wordmark">KPMG</div>
  <div class="kpmg-app-title">Capital Analysis Statement Generator</div>
</div>
<div class="kpmg-accent-strip"></div>
""", unsafe_allow_html=True)


# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_pe, tab_hf = st.tabs(["Private Equity", "Hedge Fund"])


# ═══════════════════════════ PRIVATE EQUITY TAB ═══════════════════════════════
with tab_pe:

    # Metric cards (4-up HTML grid, always visible)
    if df_latest is not None:
        _p   = partnership[:24] + ("…" if len(partnership) > 24 else "")
        _inv = str(n_investors)
        _row = str(len(df_raw))
        _aof = fmt_date(df_latest["TO_DATE"].max()) if n_investors else "—"
        _sub = "latest period"
        _sub3 = f"across {n_periods} period(s)"
    else:
        _p = _inv = _row = _aof = "—"
        _sub = "Awaiting upload"
        _sub3 = "—"

    st.markdown(f"""
    <div class="metric-grid">
      <div class="metric-card">
        <div class="label">Partnership</div>
        <div class="value">{_p}</div>
        <div class="sub">from file</div>
      </div>
      <div class="metric-card">
        <div class="label">Investors</div>
        <div class="value">{_inv}</div>
        <div class="sub">unique names</div>
      </div>
      <div class="metric-card">
        <div class="label">Data Rows</div>
        <div class="value">{_row}</div>
        <div class="sub">{_sub3}</div>
      </div>
      <div class="metric-card">
        <div class="label">As Of</div>
        <div class="value">{_aof}</div>
        <div class="sub">{_sub}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── No file state ─────────────────────────────────────────────────────────
    if uploaded_file is None:
        st.markdown("<div class='section-header'>How It Works</div>", unsafe_allow_html=True)
        st.markdown(
            "<div class='content-card'>"
            "Upload your investor Excel file using the sidebar. The tool will:<br><br>"
            "1. Detect all unique investors and the latest reporting period for each<br>"
            "2. Generate one Capital Analysis Statement per investor — Word (.docx) and PDF (.pdf)<br>"
            "3. Build a summary downloadable as Excel (.xlsx) or CSV (.csv)<br>"
            "4. Enable AI PE Insights and PE Chat powered by Gemini"
            "</div>",
            unsafe_allow_html=True,
        )

    elif read_error:
        st.error(f"Could not read file: {read_error}")

    elif missing:
        st.error(f"Missing columns: `{'`, `'.join(sorted(missing))}`")

    else:
        # ── Data preview ──────────────────────────────────────────────────────
        with st.expander("Preview source data", expanded=False):
            st.dataframe(
                df_latest[[
                    "INVESTOR_ID", "INVESTOR_NAME", "CURRENCY_CODE", "FROM_DATE", "TO_DATE",
                    "COMMITTED_CAPITAL", "OPENING_YTD_NAV", "CLOSING_YTD_NAV", "TEV", "TEV_RATIO",
                ]].rename(columns={
                    "INVESTOR_ID": "ID", "INVESTOR_NAME": "Investor", "CURRENCY_CODE": "CCY",
                    "FROM_DATE": "Period From", "TO_DATE": "Period To",
                    "COMMITTED_CAPITAL": "Committed", "OPENING_YTD_NAV": "Opening NAV",
                    "CLOSING_YTD_NAV": "Closing NAV", "TEV": "TEV", "TEV_RATIO": "Multiple",
                }),
                use_container_width=True, hide_index=True,
            )

        # Investor pills
        if chosen:
            st.markdown(
                "".join(f"<span class='investor-pill'>{inv}</span>" for inv in chosen),
                unsafe_allow_html=True,
            )

        # ── Generate ──────────────────────────────────────────────────────────
        st.markdown("<div class='section-header'>Generate Documents</div>", unsafe_allow_html=True)

        if not chosen:
            st.warning("Select at least one investor from the sidebar.")
        else:
            if st.button(f"Generate {len(chosen)} statement(s)", key="pe_generate_btn"):
                df_wrangled, wrangler_events = wrangle(df_raw)
                df_latest_w = (
                    df_wrangled.sort_values("TO_DATE", ascending=True)
                    .groupby("INVESTOR_NAME", as_index=False)
                    .last()
                )
                df_selected = df_latest_w[df_latest_w["INVESTOR_NAME"].isin(chosen)].copy()
                run_id = start_run(uploaded_file.name if uploaded_file else "unknown", len(df_latest_w))
                for ev in wrangler_events:
                    log_event(run_id, "wrangler_change", ev["investor"], ev)

                progress_bar = st.progress(0, text="Starting…")
                results_ph   = st.empty()

                docs_in_memory:      dict[str, bytes] = {}
                pdfs_in_memory:      dict[str, bytes] = {}
                result_rows:         list[dict]       = []
                validation_results:  list[dict]       = []

                for idx, (_, row) in enumerate(df_selected.iterrows()):
                    investor  = str(row["INVESTOR_NAME"]).strip()
                    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in investor)
                    progress_bar.progress(idx / len(df_selected), text=f"Generating: {investor}")

                    v_result = validate_row(row, _genai.Client(api_key=_GEMINI_KEY), run_id)
                    validation_results.append(v_result)

                    if v_result["verdict"] == "INVALID":
                        log_event(run_id, "document_failed", investor,
                                  {"reason": "validation_invalid", "notes": v_result["notes"]})
                        result_rows.append({
                            "investor": investor, "ok": False,
                            "msg": "Validation failed — " + v_result["notes"],
                        })
                        continue

                    if v_result["verdict"] == "REVALUABLE":
                        row = row.copy()
                        for field, val in v_result["corrections"].items():
                            row[field] = val
                        log_event(run_id, "validation_revalued", investor, v_result)

                    if v_result["verdict"] == "ALL_PASS":
                        log_event(run_id, "validation_pass", investor, v_result)

                    try:
                        doc = build_document(row)
                        buf = io.BytesIO()
                        doc.save(buf)
                        buf.seek(0)
                        fname = f"{safe_name}_capital_statement.docx"
                        docs_in_memory[fname] = buf.read()
                        result_rows.append({"investor": investor, "ok": True, "msg": fname})
                    except Exception as exc:
                        result_rows.append({"investor": investor, "ok": False, "msg": str(exc)})

                    try:
                        pdfs_in_memory[f"{safe_name}_capital_statement.pdf"] = build_pdf_document(row)
                    except Exception:
                        pass

                progress_bar.progress(1.0, text="Done.")

                html_rows = ""
                for r in result_rows:
                    cls = "result-ok" if r["ok"] else "result-err"
                    ico = "✓" if r["ok"] else "✗"
                    html_rows += (
                        f"<div class='result-row {cls}'>"
                        f"<span>{ico} {r['investor']}</span>"
                        f"<span style='opacity:0.7;font-size:12px;'>{r['msg']}</span>"
                        f"</div>"
                    )
                results_ph.markdown(html_rows, unsafe_allow_html=True)

                with st.expander("Validation & Audit Report"):
                    _audit_rows = [
                        {
                            "Investor":            v["investor"],
                            "Verdict":             v["verdict"],
                            "Notes":               v["notes"],
                            "Corrections Applied": bool(v["corrections"]),
                        }
                        for v in validation_results
                    ]
                    st.dataframe(pd.DataFrame(_audit_rows), use_container_width=True, hide_index=True)
                    _audit_log_path = Path("audit_log.jsonl")
                    if _audit_log_path.exists():
                        st.download_button(
                            "Download Audit Log (JSONL)",
                            data=_audit_log_path.read_bytes(),
                            file_name="audit_log.jsonl",
                            mime="application/json",
                        )
                    _audit_plain_path = Path("audit_log.txt")
                    if _audit_plain_path.exists():
                        st.download_button(
                            "Download Audit Log (Plain English)",
                            data=_audit_plain_path.read_bytes(),
                            file_name="audit_log.txt",
                            mime="text/plain",
                        )

                _success_count = sum(1 for r in result_rows if r["ok"])
                _fail_count    = sum(1 for r in result_rows if not r["ok"])
                close_run(run_id, _success_count, _fail_count)

                if docs_in_memory:
                    summary_df = build_summary_excel(df_raw, df_selected)
                    excel_buf  = io.BytesIO()
                    with pd.ExcelWriter(excel_buf, engine="openpyxl", date_format="YYYY-MM-DD") as w:
                        summary_df.to_excel(w, index=False, sheet_name="CapitalStatements")
                    excel_buf.seek(0)
                    csv_bytes = summary_df.to_csv(index=False).encode("utf-8")

                    word_zip = io.BytesIO()
                    with zipfile.ZipFile(word_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                        for fn, d in docs_in_memory.items():
                            zf.writestr(fn, d)
                    word_zip.seek(0)

                    pdf_zip = io.BytesIO()
                    with zipfile.ZipFile(pdf_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                        for fn, d in pdfs_in_memory.items():
                            zf.writestr(fn, d)
                    pdf_zip.seek(0)

                    st.session_state["gen"] = {
                        "docs":        docs_in_memory,
                        "pdfs":        pdfs_in_memory,
                        "summary_df":  summary_df,
                        "excel":       excel_buf.getvalue(),
                        "csv":         csv_bytes,
                        "word_zip":    word_zip.getvalue(),
                        "pdf_zip":     pdf_zip.getvalue(),
                        "result_rows": result_rows,
                        "df_selected": df_selected,
                        "partnership": partnership,
                        "chosen":      chosen,
                    }
                    st.session_state["pe_insights"]       = None
                    st.session_state["pe_insights_label"] = ""
                else:
                    st.error("No documents were generated.")

        # ── Post-generation (persists via session state) ───────────────────────
        gen = st.session_state.get("gen")
        if gen is not None:

            # Download panel
            st.markdown("""
            <div class="download-panel">
              <div class="download-panel-title">Download Generated Documents</div>
            </div>
            """, unsafe_allow_html=True)

            dl1, dl2, dl3, dl4 = st.columns(4)
            with dl1:
                st.download_button(
                    f"Word ZIP ({len(gen['docs'])})", data=gen["word_zip"],
                    file_name="capital_statements_word.zip",
                    mime="application/zip", key="dl_word_zip",
                )
            with dl2:
                st.download_button(
                    f"PDF ZIP ({len(gen['pdfs'])})", data=gen["pdf_zip"],
                    file_name="capital_statements_pdf.zip",
                    mime="application/zip", key="dl_pdf_zip",
                )
            with dl3:
                st.download_button(
                    "Summary Excel", data=gen["excel"],
                    file_name="capital_statements_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel",
                )
            with dl4:
                st.download_button(
                    "Summary CSV", data=gen["csv"],
                    file_name="capital_statements_summary.csv",
                    mime="text/csv", key="dl_csv",
                )

            with st.expander("Individual documents"):
                for fname, word_data in gen["docs"].items():
                    label     = fname.replace("_capital_statement.docx", "").replace("_", " ")
                    pdf_fname = fname.replace(".docx", ".pdf")
                    ca, cb    = st.columns(2)
                    with ca:
                        st.download_button(
                            f"{label} (.docx)", data=word_data, file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_word_{fname}",
                        )
                    with cb:
                        if pdf_fname in gen["pdfs"]:
                            st.download_button(
                                f"{label} (.pdf)", data=gen["pdfs"][pdf_fname],
                                file_name=pdf_fname, mime="application/pdf",
                                key=f"dl_pdf_{fname}",
                            )

            # Output preview
            st.markdown("<div class='section-header'>Output Preview</div>", unsafe_allow_html=True)
            st.dataframe(gen["summary_df"], use_container_width=True, hide_index=True)

            # AI PE Insights
            st.markdown("<div class='section-header'>AI PE Insights</div>", unsafe_allow_html=True)
            st.markdown("""
            <div class="ai-header">
              <div style="flex:1;">
                <div class="ai-title">Private Equity Investment Intelligence</div>
                <div class="ai-sub">Powered by Gemini &middot; Tailored for HNIs, Family Offices &amp; Institutional Investors</div>
              </div>
              <div class="ai-badge">PE Advisory</div>
            </div>
            """, unsafe_allow_html=True)

            ai_c1, ai_c2 = st.columns([1, 2])
            with ai_c1:
                scope = st.radio(
                    "Analysis scope:",
                    ["Portfolio Overview", "Individual Investor"],
                    key="ai_scope",
                )
            with ai_c2:
                insight_type = st.selectbox(
                    "Insight type:",
                    list(_INSIGHT_INSTRUCTIONS.keys()),
                    key="ai_insight_type",
                )

            investor_for_insight = None
            if scope == "Individual Investor":
                investor_for_insight = st.selectbox(
                    "Select investor:", gen["chosen"], key="ai_investor",
                )

            if st.button("Generate PE Insights", key="ai_generate_btn"):
                with st.spinner("Analyzing with Gemini — Private Equity Intelligence…"):
                    try:
                        df_sel = gen["df_selected"]
                        if scope == "Portfolio Overview":
                            prompt = _build_portfolio_prompt(
                                df_sel, gen["partnership"],
                                fmt_date(df_sel["TO_DATE"].max()),
                                insight_type,
                            )
                        else:
                            inv_row = df_sel[df_sel["INVESTOR_NAME"] == investor_for_insight].iloc[0]
                            prompt  = _build_investor_prompt(inv_row, insight_type)

                        st.session_state["pe_insights"] = _call_gemini(prompt)
                        st.session_state["pe_insights_label"] = (
                            f"{insight_type} — "
                            f"{'Portfolio' if scope == 'Portfolio Overview' else investor_for_insight}"
                        )
                    except Exception as e:
                        st.error(f"AI analysis failed: {e}")

            if st.session_state["pe_insights"]:
                st.markdown(
                    f"<div style='font-size:11px;font-family:\"KPMG Bold\",\"Arial Black\",sans-serif;"
                    f"text-transform:uppercase;letter-spacing:0.08em;color:#00338D;"
                    f"margin-top:16px;margin-bottom:6px;'>"
                    f"Report: {st.session_state['pe_insights_label']}</div>",
                    unsafe_allow_html=True,
                )
                st.markdown(st.session_state["pe_insights"])


# ═══════════════════════════ HEDGE FUND TAB ═══════════════════════════════════
with tab_hf:

    st.markdown(
        "<div class='section-header'>Hedge Fund Capital Statement Generator</div>",
        unsafe_allow_html=True,
    )

    # ── PCAP upload ────────────────────────────────────────────────────────────
    hf_file = st.file_uploader(
        "Upload KPMG PCAP (.xlsx)",
        type=["xlsx"],
        key="hf_upload",
        help=(
            "Pre-calculated PCAP Excel with one row per investor and ~142 columns. "
            "Must contain investor names and capital account balances (CQ/YTD/ITD)."
        ),
    )

    # ── Optional CF Ledger upload ──────────────────────────────────────────────
    cf_ledger_file = st.file_uploader(
        "Upload CF Ledger (.xlsx)  — optional, enables CF_Aggregator & Cashflow_IRR sheets",
        type=["xlsx"],
        key="hf_cf_ledger_upload",
        help=(
            "Transaction-level cashflow data. Required columns: INVESTOR_NAME, TRANSACTION_DATE, TYPE, AMOUNT. "
            "TYPE values: Contribution | Distribution | DRIP | Transfer | Redemption. "
            "Optional: TRANSACTION_ID, QUARTER, SUB_TYPE (In/Out for Transfer), UNITS, NOTES."
        ),
    )

    hf_pcap_df    = None
    hf_read_error = None
    cf_ledger_df  = None

    if hf_file is not None:
        try:
            hf_file.seek(0)
            hf_pcap_df = read_hf_pcap_from_upload(hf_file)
        except Exception as exc:
            hf_read_error = str(exc)

    if cf_ledger_file is not None:
        try:
            cf_ledger_file.seek(0)
            cf_ledger_df = pd.read_excel(cf_ledger_file)
        except Exception:
            cf_ledger_df = None

    # ── Metric cards ───────────────────────────────────────────────────────────
    _hf_investors = "—"
    _hf_nav_cq    = "—"
    _hf_net_irr   = "—"
    _hf_tvpi      = "—"

    if hf_pcap_df is not None:
        _hf_investors = str(len(hf_pcap_df))
        _nav = hf_pcap_df["END_CAP_CQ"].dropna().sum() if "END_CAP_CQ" in hf_pcap_df.columns else 0.0
        _irr = hf_pcap_df["NET_IRR"].dropna().mean()   if "NET_IRR"    in hf_pcap_df.columns else None
        _tv  = hf_pcap_df["TVPI"].dropna().mean()      if "TVPI"       in hf_pcap_df.columns else None
        _hf_nav_cq  = fmt_usd(_nav)
        _hf_net_irr = f"{_irr:.2f}%" if _irr is not None else "—"
        _hf_tvpi    = f"{_tv:.2f}x"  if _tv  is not None else "—"

    st.markdown(f"""
    <div class="metric-grid">
      <div class="metric-card">
        <div class="label">Investors</div>
        <div class="value">{_hf_investors}</div>
        <div class="sub">unique LPs</div>
      </div>
      <div class="metric-card">
        <div class="label">Total NAV (CQ)</div>
        <div class="value" style="font-size:16px;">{_hf_nav_cq}</div>
        <div class="sub">ending partner's capital</div>
      </div>
      <div class="metric-card">
        <div class="label">Avg Net IRR</div>
        <div class="value">{_hf_net_irr}</div>
        <div class="sub">across all LPs</div>
      </div>
      <div class="metric-card">
        <div class="label">Avg TVPI</div>
        <div class="value">{_hf_tvpi}</div>
        <div class="sub">total value multiple</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Error states ───────────────────────────────────────────────────────────
    if hf_file is None:
        st.markdown("""
        <div class="hf-banner">
          <b>Upload the KPMG Hedge Fund PCAP Excel to begin.</b><br>
          The file should contain one row per investor with pre-calculated capital account data.<br>
          Required: <code>Investor Name - Legal Name from Master List</code> and
          <code>Ending Partner's Capital CQ</code> columns (plus any subset of the 140+ PCAP columns).
        </div>
        """, unsafe_allow_html=True)

    elif hf_read_error:
        st.error(f"Could not read PCAP file: {hf_read_error}")

    else:
        # ── Data preview ───────────────────────────────────────────────────────
        with st.expander("Preview PCAP Data", expanded=False):
            st.dataframe(hf_pcap_df, use_container_width=True, hide_index=True)

        # ── Investor selection ─────────────────────────────────────────────────
        hf_all_investors = sorted(hf_pcap_df["INVESTOR_NAME"].dropna().unique().tolist())
        hf_chosen = st.multiselect(
            "Select investors to generate:",
            options=hf_all_investors,
            default=hf_all_investors,
            key="hf_sel_investors",
        )
        if hf_chosen:
            st.markdown(
                "".join(f"<span class='investor-pill'>{inv}</span>" for inv in hf_chosen),
                unsafe_allow_html=True,
            )

        # ── Generate ───────────────────────────────────────────────────────────
        st.markdown("<div class='section-header'>Generate Documents</div>", unsafe_allow_html=True)

        if not hf_chosen:
            st.warning("Select at least one investor above.")
        else:
            if st.button(f"Generate {len(hf_chosen)} HF Statement(s)", key="hf_generate_btn"):
                pcap_df = hf_pcap_df[hf_pcap_df["INVESTOR_NAME"].isin(hf_chosen)].copy()

                if not pcap_df.empty:
                    hf_docs, hf_pdfs = {}, {}
                    result_rows      = []
                    progress_bar     = st.progress(0, text="Generating statements…")

                    for idx, (_, prow) in enumerate(pcap_df.iterrows()):
                        investor  = str(prow["INVESTOR_NAME"]).strip()
                        safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in investor)
                        progress_bar.progress(idx / len(pcap_df), text=f"Generating: {investor}")

                        # Word document
                        try:
                            doc_obj  = build_hf_docx(prow)
                            _buf     = __import__("io").BytesIO()
                            doc_obj.save(_buf); _buf.seek(0)
                            fname = f"{safe_name}_hf_capital_statement.docx"
                            hf_docs[fname] = _buf.read()
                            result_rows.append({"investor": investor, "ok": True, "msg": fname})
                        except Exception as exc:
                            result_rows.append({"investor": investor, "ok": False, "msg": str(exc)})

                        # PDF
                        try:
                            hf_pdfs[f"{safe_name}_hf_capital_statement.pdf"] = build_hf_pdf(prow)
                        except Exception:
                            pass

                    progress_bar.progress(1.0, text="Done.")

                    # Result rows display
                    _html = ""
                    for rr in result_rows:
                        cls = "result-ok" if rr["ok"] else "result-err"
                        ico = "✓" if rr["ok"] else "✗"
                        _html += (
                            f"<div class='result-row {cls}'>"
                            f"<span>{ico} {rr['investor']}</span>"
                            f"<span style='opacity:0.7;font-size:12px;'>{rr['msg']}</span>"
                            f"</div>"
                        )
                    st.markdown(_html, unsafe_allow_html=True)

                    # Build ZIPs and companion Excel
                    import io as _io, zipfile as _zf

                    word_zip = _io.BytesIO()
                    with _zf.ZipFile(word_zip, "w", _zf.ZIP_DEFLATED) as zf:
                        for fn, d in hf_docs.items(): zf.writestr(fn, d)
                    word_zip.seek(0)

                    pdf_zip = _io.BytesIO()
                    with _zf.ZipFile(pdf_zip, "w", _zf.ZIP_DEFLATED) as zf:
                        for fn, d in hf_pdfs.items(): zf.writestr(fn, d)
                    pdf_zip.seek(0)

                    try:
                        companion_xlsx = build_hf_workbook(pcap_df, cf_ledger_df)
                    except Exception:
                        companion_xlsx = None

                    # Summary Excel
                    summary_cols = [
                        "INVESTOR_NAME", "INCEPTION_DATE", "REPT_CCY",
                        "BEG_CAP_CQ", "END_CAP_CQ", "BEG_UNITS_CQ", "END_UNITS_CQ",
                        "BEG_CAP_YTD", "END_CAP_YTD", "BEG_CAP_ITD", "END_CAP_ITD",
                        "BEG_PX", "END_PX", "TOTAL_COMMIT", "FUNDED_COMMIT", "AVAIL_COMMIT",
                        "GROSS_IRR", "NET_IRR", "DPI", "RVPI", "TVPI",
                        "HURDLE_AMT_ITD", "EXCESS_HURDLE", "GP_CATCHUP_AMT", "LP_NET_WF",
                        "MGMT_FEE_RATE", "INC_FEE_RATE", "PREF_RET",
                    ]
                    summary_df = pcap_df[[c for c in summary_cols if c in pcap_df.columns]].copy()
                    summ_buf = _io.BytesIO()
                    with pd.ExcelWriter(summ_buf, engine="openpyxl", date_format="YYYY-MM-DD") as w:
                        summary_df.to_excel(w, index=False, sheet_name="HF_Summary")
                    summ_buf.seek(0)

                    st.session_state["hf_gen"] = {
                        "docs":           hf_docs,
                        "pdfs":           hf_pdfs,
                        "pcap_df":        pcap_df,
                        "word_zip":       word_zip.getvalue(),
                        "pdf_zip":        pdf_zip.getvalue(),
                        "summary_excel":  summ_buf.getvalue(),
                        "companion_xlsx": companion_xlsx,
                        "result_rows":    result_rows,
                        "chosen":         hf_chosen,
                    }
                    st.session_state["hf_insights"]       = None
                    st.session_state["hf_insights_label"] = ""

        # ── Post-generation downloads + preview ────────────────────────────────
        hf_gen = st.session_state.get("hf_gen")
        if hf_gen is not None:

            st.markdown("""
            <div class="download-panel">
              <div class="download-panel-title">Download Generated Documents</div>
            </div>
            """, unsafe_allow_html=True)

            _dl1, _dl2, _dl3, _dl4 = st.columns(4)
            with _dl1:
                st.download_button(
                    f"Word ZIP ({len(hf_gen['docs'])})", data=hf_gen["word_zip"],
                    file_name="hf_capital_statements_word.zip",
                    mime="application/zip", key="hf_dl_word",
                )
            with _dl2:
                st.download_button(
                    f"PDF ZIP ({len(hf_gen['pdfs'])})", data=hf_gen["pdf_zip"],
                    file_name="hf_capital_statements_pdf.zip",
                    mime="application/zip", key="hf_dl_pdf",
                )
            with _dl3:
                st.download_button(
                    "Summary Excel", data=hf_gen["summary_excel"],
                    file_name="hf_capital_statements_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="hf_dl_summary",
                )
            with _dl4:
                if hf_gen["companion_xlsx"]:
                    st.download_button(
                        "PCAP Model (.xlsx)", data=hf_gen["companion_xlsx"],
                        file_name="hf_pcap_model.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="hf_dl_companion",
                    )

            with st.expander("Individual documents"):
                for fname, word_data in hf_gen["docs"].items():
                    label     = fname.replace("_hf_capital_statement.docx", "").replace("_", " ")
                    pdf_fname = fname.replace(".docx", ".pdf")
                    _ca, _cb  = st.columns(2)
                    with _ca:
                        st.download_button(
                            f"{label} (.docx)", data=word_data, file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"hf_dl_word_{fname}",
                        )
                    with _cb:
                        if pdf_fname in hf_gen["pdfs"]:
                            st.download_button(
                                f"{label} (.pdf)", data=hf_gen["pdfs"][pdf_fname],
                                file_name=pdf_fname, mime="application/pdf",
                                key=f"hf_dl_pdf_{fname}",
                            )

            # PCAP preview table
            st.markdown("<div class='section-header'>PCAP Output Preview</div>", unsafe_allow_html=True)
            _preview_cols = [
                "INVESTOR_NAME", "INCEPTION_DATE", "REPT_CCY",
                "BEG_CAP_CQ", "END_CAP_CQ", "BEG_PX", "END_PX",
                "GROSS_IRR", "NET_IRR", "TVPI", "LP_NET_WF",
            ]
            _preview_df = hf_gen["pcap_df"][[c for c in _preview_cols if c in hf_gen["pcap_df"].columns]].copy()
            st.dataframe(_preview_df, use_container_width=True, hide_index=True)

            # ── AI HF Insights ─────────────────────────────────────────────────
            st.markdown("<div class='section-header'>AI HF Insights</div>", unsafe_allow_html=True)
            st.markdown("""
            <div class="ai-header">
              <div style="flex:1;">
                <div class="ai-title">Hedge Fund Investment Intelligence</div>
                <div class="ai-sub">Powered by Gemini &middot; PCAP Analysis, Waterfall, Stress Testing &amp; IRR Attribution</div>
              </div>
              <div class="ai-badge">HF Advisory</div>
            </div>
            """, unsafe_allow_html=True)

            _hf_ai_c1, _hf_ai_c2 = st.columns([1, 2])
            with _hf_ai_c1:
                hf_scope = st.radio(
                    "Analysis scope:",
                    ["Portfolio Overview", "Individual Investor"],
                    key="hf_ai_scope",
                )
            with _hf_ai_c2:
                hf_insight_type = st.selectbox(
                    "Insight type:",
                    list(_HF_INSIGHT_INSTRUCTIONS.keys()),
                    key="hf_ai_insight_type",
                )

            hf_investor_for_insight = None
            if hf_scope == "Individual Investor":
                hf_investor_for_insight = st.selectbox(
                    "Select investor:", hf_gen["chosen"], key="hf_ai_investor",
                )

            if st.button("Generate HF Insights", key="hf_ai_generate_btn"):
                with st.spinner("Analyzing with Gemini — Hedge Fund Intelligence…"):
                    try:
                        if hf_scope == "Portfolio Overview":
                            hf_prompt = _build_hf_portfolio_prompt(
                                hf_gen["pcap_df"], hf_insight_type,
                            )
                        else:
                            inv_row = hf_gen["pcap_df"][
                                hf_gen["pcap_df"]["INVESTOR_NAME"] == hf_investor_for_insight
                            ].iloc[0]
                            hf_prompt = _build_hf_investor_prompt(
                                inv_row, hf_insight_type,
                            )
                        st.session_state["hf_insights"] = _call_gemini(hf_prompt)
                        st.session_state["hf_insights_label"] = (
                            f"{hf_insight_type} — "
                            f"{'Portfolio' if hf_scope == 'Portfolio Overview' else hf_investor_for_insight}"
                        )
                    except Exception as e:
                        st.error(f"AI analysis failed: {e}")

            if st.session_state["hf_insights"]:
                st.markdown(
                    f"<div style='font-size:11px;font-family:\"KPMG Bold\",\"Arial Black\",sans-serif;"
                    f"text-transform:uppercase;letter-spacing:0.08em;color:#00338D;"
                    f"margin-top:16px;margin-bottom:6px;'>"
                    f"Report: {st.session_state['hf_insights_label']}</div>",
                    unsafe_allow_html=True,
                )
                st.markdown(st.session_state["hf_insights"])
