"""
KPMG Capital Analysis Statement Generator — FastAPI Backend
Run: uvicorn api:app --reload --port 8000

Wraps the existing Python pipeline modules (never modifies generate_capital_statements.py).
Exposes REST endpoints consumed by frontend/index.html.
"""

from __future__ import annotations

import io
import json
import os
import time
import uuid
import zipfile
from pathlib import Path
from typing import AsyncGenerator

import pandas as pd
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, StreamingResponse
from pydantic import BaseModel

load_dotenv()

# ── ReportLab (PDF builder) ───────────────────────────────────────────────────
from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ── KPMG pipeline modules ─────────────────────────────────────────────────────
from generate_capital_statements import (
    REQUIRED_COLS,
    build_document,
    build_summary_excel,
    fmt_date,
    fmt_ratio,
    fmt_usd,
)
from data_wrangler import wrangle
from audit_trail import close_run, log_event, start_run
from validation_agent import validate_row
from hf_pcap_engine import HF_REQUIRED_COLS, read_hf_pcap_from_upload
from hf_statement_generator import build_hf_docx, build_hf_pdf
from hf_excel_generator import build_hf_workbook

# ── Gemini ────────────────────────────────────────────────────────────────────
from google import genai as _genai

_GEMINI_KEY   = os.environ.get("GEMINI_API_KEY", "")
_GEMINI_MODEL = "models/gemini-2.5-pro"
_MAX_TOKENS   = 65536


# ── ReportLab paragraph styles (module-level to avoid registry collisions) ────
_BASE  = getSampleStyleSheet()
_RED_C = rl_colors.HexColor("#D73B3E")

_PS_CONF     = ParagraphStyle("api_kpmg_conf",     parent=_BASE["Normal"], fontSize=9,  textColor=_RED_C,  alignment=TA_RIGHT, fontName="Helvetica-Bold")
_PS_TITLE    = ParagraphStyle("api_kpmg_title",    parent=_BASE["Normal"], fontSize=16, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=2)
_PS_SUBTITLE = ParagraphStyle("api_kpmg_subtitle", parent=_BASE["Normal"], fontSize=12, alignment=TA_CENTER, spaceAfter=4)
_PS_NORMAL   = ParagraphStyle("api_kpmg_normal",   parent=_BASE["Normal"], fontSize=11, leading=15)
_PS_INDENT   = ParagraphStyle("api_kpmg_indent",   parent=_BASE["Normal"], fontSize=11, leading=15, leftIndent=14)
_PS_VALUE    = ParagraphStyle("api_kpmg_value",    parent=_BASE["Normal"], fontSize=11, leading=15, alignment=TA_RIGHT)
_PS_FOOT     = ParagraphStyle("api_kpmg_foot",     parent=_BASE["Normal"], fontSize=9,  fontName="Helvetica-Oblique")


# ── Gemini personas ───────────────────────────────────────────────────────────
_PE_SYSTEM = """\
You are a senior Private Equity investment advisor with 20+ years of experience. You specialise in analyzing LP capital accounts and fund performance to \
provide strategic insights for HNIs, family offices, and institutional investors.

Your expertise covers LP/GP dynamics, DPI/RVPI/TVPI, fee analysis, distribution planning, capital deployment \
pacing, exit strategy, and risk-adjusted return benchmarking (top-quartile PE: TVPI >2.0x; median ~1.6x).

Be direct, data-driven, and structured. Use PE terminology. Flag red flags clearly. \
Every insight must be tied to the specific numbers provided — no generic commentary.\
"""

_HF_SYSTEM = """\
You are a senior Hedge Fund analyst with 20+ years of experience. You specialise in hedge fund LP capital accounts, \
PCAP analysis, unit-based NAV attribution, and risk-adjusted returns.

Your expertise covers unit pricing, NAV attribution, IRR vs. TWR, management/incentive fee drag, \
waterfall mechanics, stress testing, and LP redemption risk.

Be direct, data-driven, and structured. Use hedge fund terminology precisely. \
Every insight must be tied to the specific numbers provided — no generic commentary.\
"""

_PE_INSIGHT_INSTRUCTIONS = {
    "full":      "Cover: (1) Performance vs PE benchmarks (TVPI/DPI/RVPI), (2) Capital account health, (3) Fee drag on net returns, (4) YTD activity, (5) Key risks and opportunities, (6) Specific recommendations.",
    "fee":       "Cover: (1) Management fee as % of committed/contributed vs industry norm (1.5–2%), (2) Incentive fee vs value created, (3) Total fee drag on gross-to-net return, (4) Fee-adjusted DPI and TVPI, (5) Recommendations.",
    "liquidity": "Cover: (1) DPI and capital return timeline, (2) RVPI realization outlook, (3) YTD distribution pace, (4) Liquidity needs and secondary market optionality, (5) Distribution strategy recommendation.",
    "exit":      "Cover: (1) TVPI vs typical PE exit multiples, (2) Unrealized gain and exit readiness, (3) NAV trajectory signals, (4) Optimal hold period, (5) Exit scenario modeling (base/bull/bear), (6) Timing recommendation.",
    "pacing":    "Cover: (1) Utilization rate vs typical PE pace (60–80% by year 3–4), (2) Unfunded commitment and remaining call risk, (3) Over-commitment risk, (4) Capital call timing projections, (5) Recommendations for managing remaining obligations.",
}

_HF_INSIGHT_INSTRUCTIONS = {
    "full":      "Cover: (1) NAV and unit price movement attribution, (2) IRR vs. hurdle, (3) Fee drag (management + incentive), (4) Capital account composition, (5) Stress test implications, (6) Key risks and strategic recommendations.",
    "fee":       "Cover: (1) Incentive fee as % of gross return, (2) Hurdle rate adequacy, (3) GP carry vs. LP net return split, (4) Fee-adjusted IRR, (5) Comparison to industry norms (2-and-20 vs actual structure), (6) Recommendations.",
    "waterfall": "Cover: (1) Excess return above hurdle, (2) GP vs LP distribution split, (3) Effective carry rate, (4) LP net return after carry, (5) Comparison to contributed capital, (6) Pacing of distributions.",
    "stress":    "Cover: (1) NAV sensitivity at -10%/-20%/-30% haircuts, (2) IRR impact under each scenario, (3) Break-even analysis, (4) Probability-weighted return, (5) Liquidity and redemption risk, (6) Hedging / risk mitigation recommendations.",
    "nav":       "Cover: (1) Unit price change decomposition (income / unrealized / realized / expenses / fees), (2) Investor share vs. fund-level P&L, (3) DRIP impact on units and NAV, (4) Redemption dilution effect, (5) Ending NAV quality assessment.",
}


# ── In-memory session storage ─────────────────────────────────────────────────
_pe_sessions:  dict[str, dict] = {}   # token → {df_wrangled, df_latest, filename}
_hf_sessions:  dict[str, dict] = {}   # token → {pcap_df, filename}
_cf_sessions:  dict[str, dict] = {}   # token → {cf_df}
_pe_results:   dict[str, dict] = {}   # run_id → {word_zip, pdf_zip, summary_excel, summary_csv}
_hf_results:   dict[str, dict] = {}   # run_id → {word_zip, pdf_zip, summary_excel, companion_xlsx}


# ── Helpers ───────────────────────────────────────────────────────────────────
def _df_to_records(df: pd.DataFrame) -> list[dict]:
    """Serialize a DataFrame to JSON-safe records for stateless clients."""
    records = []
    for _, row in df.iterrows():
        rec = {}
        for k, v in row.items():
            if isinstance(v, float) and pd.isna(v):
                rec[k] = None
            elif hasattr(v, "item"):
                rec[k] = v.item()
            elif hasattr(v, "isoformat"):
                rec[k] = str(v)[:10]
            else:
                rec[k] = v
        records.append(rec)
    return records


def _records_to_df(records: list[dict]) -> pd.DataFrame:
    """Reconstruct a DataFrame from JSON records, restoring date columns."""
    df = pd.DataFrame(records)
    for col in ["TO_DATE", "FROM_DATE", "INCEPTION_DATE"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def _safe_div(num, den) -> float:
    try:
        return float(num) / float(den) if float(den) != 0 else 0.0
    except (TypeError, ValueError):
        return 0.0


def _row_to_dict(row: pd.Series) -> dict:
    """Convert a DataFrame row to a JSON-serialisable dict."""
    d = {}
    for k, v in row.items():
        if pd.isna(v) if isinstance(v, float) else False:
            d[k] = None
        elif hasattr(v, "item"):  # numpy scalar
            d[k] = v.item()
        elif hasattr(v, "isoformat"):  # date/datetime
            d[k] = str(v)[:10]
        else:
            d[k] = v
    return d


# ── PDF builder (extracted from app.py) ───────────────────────────────────────
def build_pdf_document(row: pd.Series) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch,
        leftMargin=1.0 * inch, rightMargin=1.0 * inch,
    )
    COL_W = [4.25 * inch, 2.25 * inch]

    def _two_col(label: str, value: str, indented: bool = False) -> Table:
        lbl_s = _PS_INDENT if indented else _PS_NORMAL
        data  = [[Paragraph(label, lbl_s), Paragraph(value, _PS_VALUE)]]
        t = Table(data, colWidths=COL_W)
        t.setStyle(TableStyle([
            ("TOPPADDING",    (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING",   (0, 0), (-1, -1), 0),
            ("RIGHTPADDING",  (-1, 0), (-1, -1), 0),
        ]))
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

    story += [_section("Summary of Capital Account"), sp(6)]
    story.append(_two_col(f"Opening Capital balance as on {fmt_date(row['FROM_DATE'])}", fmt_usd(row["OPENING_YTD_NAV"])))
    story.append(_two_col("Capital contributions during the year",  fmt_usd(row["YTD_CONTRIBUTION"])))
    story.append(_two_col("Distributions during the year",          fmt_usd(row["YTD_DISTRIBUTION"])))
    story.append(sp(6))
    story.append(Paragraph("Net investment activity:", _PS_NORMAL))
    net_income = float(row["INVESTMENT_INCOME"]) - float(row["INVESTMENT_EXPENSE"])
    story.append(_two_col("Investment and other income",                fmt_usd(net_income),                   indented=True))
    story.append(_two_col("Net unrealized appreciation (depreciation)", fmt_usd(row["UNREALIZED_GAINS_LOSS"]), indented=True))
    story.append(_two_col("Net realized gain (loss)",                   fmt_usd(row["REALIZED_GAINS_LOSS"]),   indented=True))
    story.append(sp(6))
    story.append(_two_col("Management fees for the period", fmt_usd(row["MANAGEMENT_FEE"])))
    story.append(_two_col("Incentive fees for the period",  fmt_usd(row["INCENTIVE_FEE"])))
    story.append(sp(6))
    story.append(_two_col(f"Capital balance (remaining value) at {fmt_date(row['TO_DATE'])} *", fmt_usd(row["CLOSING_YTD_NAV"])))
    story.append(sp(16))

    story += [_section("Summary of Capital Commitment"), sp(6)]
    committed   = float(row["COMMITTED_CAPITAL"])
    contributed = float(row["INCEPTION_TO_DATE_CONTRIBUTION"])
    story.append(_two_col("Capital commitment per subscription agreement (A)", fmt_usd(committed)))
    story.append(_two_col("Capital contributed to date (B)",                   fmt_usd(contributed)))
    story.append(_two_col("Remaining capital commitment (A-B)",                fmt_usd(committed - contributed)))
    story.append(sp(16))

    story += [_section("Summary of Distributions and Valuation"), sp(6)]
    story.append(_two_col("Total capital contributed to date",                                  fmt_usd(row["INCEPTION_TO_DATE_CONTRIBUTION"])))
    story.append(_two_col("Total distributions to date",                                         fmt_usd(row["INCEPTION_TO_DATE_DISTRIBUTION"])))
    story.append(_two_col(f"Capital balance (remaining value) at {fmt_date(row['TO_DATE'])} *", fmt_usd(row["CLOSING_YTD_NAV"])))
    story.append(_two_col("Total Estimated Value (distributions + balance)",                     fmt_usd(row["TEV"])))
    story.append(_two_col("Total Estimated Value as net multiple",                               fmt_ratio(row["TEV_RATIO"])))
    story.append(sp(20))
    story.append(Paragraph(
        "* Represents remaining value. The remaining value is based upon available "
        "information and may not represent amounts which might ultimately be realized.",
        _PS_FOOT,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── Gemini prompt builders ────────────────────────────────────────────────────
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


def _build_pe_portfolio_prompt(df_sel, partnership, as_of, insight_key):
    ctx  = _portfolio_context(df_sel, partnership)
    instr = _PE_INSIGHT_INSTRUCTIONS.get(insight_key, _PE_INSIGHT_INSTRUCTIONS["full"])
    return f"AS OF: {as_of}\n\n{ctx}\n\nANALYSIS REQUEST (Portfolio Level):\n{instr}"


def _build_pe_investor_prompt(row: pd.Series, insight_key: str) -> str:
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
    instr = _PE_INSIGHT_INSTRUCTIONS.get(insight_key, _PE_INSIGHT_INSTRUCTIONS["full"])
    return f"{ctx}\n\nANALYSIS REQUEST (Investor Level):\n{instr}"


def _build_hf_portfolio_prompt(pcap_df: pd.DataFrame, insight_key: str) -> str:
    def _col(col, fn):
        if col in pcap_df.columns:
            vals = pcap_df[col].dropna()
            return fn(vals) if len(vals) else 0.0
        return 0.0

    total_nav   = _col("END_CAP_CQ",  sum)
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
        f"Avg Beg Unit Price: {fmt_usd(beg_px)}  |  Avg End Unit Price: {fmt_usd(end_px)}",
        f"Total NAV (CQ Ending): {fmt_usd(total_nav)}",
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

    instr = _HF_INSIGHT_INSTRUCTIONS.get(insight_key, _HF_INSIGHT_INSTRUCTIONS["full"])
    return f"{chr(10).join(lines)}\n\nANALYSIS REQUEST (Portfolio Level):\n{instr}"


def _build_hf_investor_prompt(row: pd.Series, insight_key: str) -> str:
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
        f"Hurdle Amt (ITD): {fmt_usd(_f('HURDLE_AMT_ITD'))}  |  Excess Over Hurdle: {fmt_usd(_f('EXCESS_HURDLE'))}\n"
        f"GP Catch-Up: {fmt_usd(_f('GP_CATCHUP_AMT'))}  |  LP Net WF Share: {fmt_usd(_f('LP_NET_WF'))}\n"
        f"Total Commit: {fmt_usd(_f('TOTAL_COMMIT'))}  |  Funded: {fmt_usd(_f('FUNDED_COMMIT'))}\n"
        f"Lock-up: {_f('LOCKUP_MO'):.0f} months  |  HWM Active: {_s('HWM_ACTIVE')}\n"
        f"Pref Return: {_f('PREF_RET'):.2f}%  |  Inc Fee Rate: {_f('INC_FEE_RATE'):.2f}%"
    )
    instr = _HF_INSIGHT_INSTRUCTIONS.get(insight_key, _HF_INSIGHT_INSTRUCTIONS["full"])
    return f"{ctx}\n\nANALYSIS REQUEST (Investor Level):\n{instr}"


_PLACEHOLDER_KEYS = {"your_key_here", "your-key-here", "YOUR_KEY_HERE", ""}


def _check_gemini_key() -> str | None:
    """Return an error string if the key is missing or a placeholder, else None."""
    if not _GEMINI_KEY or _GEMINI_KEY.strip() in _PLACEHOLDER_KEYS:
        return "GEMINI_API_KEY is not configured. Add your real key to .env and restart the server."
    return None


# ── Gemini calls ──────────────────────────────────────────────────────────────
def _call_gemini_sync(prompt: str, system: str) -> str:
    err = _check_gemini_key()
    if err:
        raise RuntimeError(err)
    client = _genai.Client(api_key=_GEMINI_KEY)
    response = client.models.generate_content(
        model=_GEMINI_MODEL,
        config=_genai.types.GenerateContentConfig(
            system_instruction=system,
            max_output_tokens=_MAX_TOKENS,
        ),
        contents=prompt,
    )
    return response.text or ""


def _call_gemini_chat_sync(messages: list[dict], system: str) -> str:
    err = _check_gemini_key()
    if err:
        return f"⚠ {err}"
    client = _genai.Client(api_key=_GEMINI_KEY)
    contents = [
        _genai.types.Content(
            role="user" if m["role"] == "user" else "model",
            parts=[_genai.types.Part(text=m["content"])],
        )
        for m in messages
    ]
    response = client.models.generate_content(
        model=_GEMINI_MODEL,
        config=_genai.types.GenerateContentConfig(
            system_instruction=system,
            max_output_tokens=_MAX_TOKENS,
        ),
        contents=contents,
    )
    return response.text or ""


# ── FastAPI app ───────────────────────────────────────────────────────────────
app = FastAPI(title="KPMG Capital Statement Generator API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # tighten to ["http://localhost:8080"] in prod
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ═══════════════════════════ PE ENDPOINTS ═════════════════════════════════════

@app.post("/api/pe/upload")
async def pe_upload(file: UploadFile = File(...)):
    """Accept PE xlsx, wrangle, return investor list + preview data."""
    content = await file.read()
    try:
        df_raw = pd.read_excel(io.BytesIO(content))
    except Exception as exc:
        return {"ok": False, "error": f"Could not read file: {exc}"}

    missing = REQUIRED_COLS - set(df_raw.columns)
    if missing:
        return {"ok": False, "missing_cols": sorted(missing), "error": "Required columns missing"}

    df_raw["TO_DATE"]   = pd.to_datetime(df_raw["TO_DATE"],   errors="coerce")
    df_raw["FROM_DATE"] = pd.to_datetime(df_raw["FROM_DATE"], errors="coerce")

    df_wrangled, _ = wrangle(df_raw)
    df_latest = (
        df_wrangled.sort_values("TO_DATE", ascending=True)
        .groupby("INVESTOR_NAME", as_index=False).last()
    )

    session_token = str(uuid.uuid4())
    _pe_sessions[session_token] = {
        "df_raw":      df_raw,
        "df_wrangled": df_wrangled,
        "df_latest":   df_latest,
        "filename":    file.filename or "upload.xlsx",
    }

    partnership = str(df_latest["PARTNERSHIP_NAME"].iloc[0]) if len(df_latest) else "—"
    as_of       = str(df_latest["TO_DATE"].max())[:10] if len(df_latest) else "—"
    n_periods   = int(df_raw["TO_DATE"].nunique())

    # Preview rows (JSON-safe)
    preview = []
    for _, row in df_latest.iterrows():
        preview.append({
            "investor_id":  str(row.get("INVESTOR_ID", "")),
            "name":         str(row["INVESTOR_NAME"]),
            "ccy":          str(row.get("CURRENCY_CODE", "USD")),
            "from_date":    str(row["FROM_DATE"])[:10],
            "to_date":      str(row["TO_DATE"])[:10],
            "committed":    float(row.get("COMMITTED_CAPITAL", 0) or 0),
            "itd_contrib":  float(row.get("INCEPTION_TO_DATE_CONTRIBUTION", 0) or 0),
            "itd_dist":     float(row.get("INCEPTION_TO_DATE_DISTRIBUTION", 0) or 0),
            "opening_nav":  float(row.get("OPENING_YTD_NAV", 0) or 0),
            "ytd_contrib":  float(row.get("YTD_CONTRIBUTION", 0) or 0),
            "ytd_dist":     float(row.get("YTD_DISTRIBUTION", 0) or 0),
            "inv_income":   float(row.get("INVESTMENT_INCOME", 0) or 0),
            "inv_expense":  float(row.get("INVESTMENT_EXPENSE", 0) or 0),
            "unreal":       float(row.get("UNREALIZED_GAINS_LOSS", 0) or 0),
            "real":         float(row.get("REALIZED_GAINS_LOSS", 0) or 0),
            "mgmt_fee":     float(row.get("MANAGEMENT_FEE", 0) or 0),
            "inc_fee":      float(row.get("INCENTIVE_FEE", 0) or 0),
            "closing_nav":  float(row.get("CLOSING_YTD_NAV", 0) or 0),
            "tev":          float(row.get("TEV", 0) or 0),
            "tev_ratio":    float(row.get("TEV_RATIO", 0) or 0),
            "verdict":      "PENDING",
        })

    return {
        "ok":            True,
        "session_token": session_token,
        "partnership":   partnership,
        "n_investors":   len(df_latest),
        "n_periods":     n_periods,
        "as_of":         as_of,
        "investors":     df_latest["INVESTOR_NAME"].tolist(),
        "preview":       preview,
        "session_data":  _df_to_records(df_latest),
        "missing_cols":  [],
        "error":         None,
    }


class PeGenerateRequest(BaseModel):
    session_token: str
    investors: list[str]
    session_data: list[dict] | None = None  # stateless fallback for serverless


@app.post("/api/pe/generate")
async def pe_generate(req: PeGenerateRequest):
    """Run wrangle→validate→generate pipeline, stream results via SSE."""
    sess = _pe_sessions.get(req.session_token)
    if sess:
        df_wrangled   = sess["df_wrangled"]
        df_latest_all = sess["df_latest"]
        filename      = sess["filename"]
        df_raw        = sess["df_raw"]
    elif req.session_data:
        df_latest_all = _records_to_df(req.session_data)
        df_wrangled   = df_latest_all
        filename      = "upload.xlsx"
        df_raw        = df_latest_all
    else:
        raise HTTPException(404, "Session not found — re-upload the file")

    df_selected = df_latest_all[df_latest_all["INVESTOR_NAME"].isin(req.investors)].copy()
    if df_selected.empty:
        raise HTTPException(400, "No matching investors in session")

    gemini_client = _genai.Client(api_key=_GEMINI_KEY) if _GEMINI_KEY else None

    async def event_stream() -> AsyncGenerator[str, None]:
        run_id = start_run(filename, len(df_selected))
        docs_mem:  dict[str, bytes] = {}
        pdfs_mem:  dict[str, bytes] = {}
        success = fail = 0

        for _, row in df_selected.iterrows():
            investor  = str(row["INVESTOR_NAME"]).strip()
            safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in investor)

            # Validation (Gemini call — blocking, run in threadpool)
            if gemini_client:
                v_result = await run_in_threadpool(validate_row, row, gemini_client, run_id)
            else:
                # No API key — local checks only (fast)
                v_result = await run_in_threadpool(validate_row, row, None, run_id)

            verdict = v_result["verdict"]

            if verdict == "INVALID":
                log_event(run_id, "document_failed", investor,
                          {"reason": "validation_invalid", "notes": v_result["notes"]})
                fail += 1
                event = {
                    "type": "progress", "investor": investor, "ok": False,
                    "verdict": "INVALID",
                    "file": None,
                    "error": "Validation failed — " + v_result["notes"],
                }
                yield f"data: {json.dumps(event)}\n\n"
                continue

            if verdict == "REVALUABLE":
                row = row.copy()
                for field, val in v_result.get("corrections", {}).items():
                    row[field] = val
                log_event(run_id, "validation_revalued", investor, v_result)

            if verdict == "ALL_PASS":
                log_event(run_id, "validation_pass", investor, v_result)

            # Document generation
            fname_docx = f"{safe_name}_capital_statement.docx"
            fname_pdf  = f"{safe_name}_capital_statement.pdf"
            try:
                doc = build_document(row)
                buf = io.BytesIO(); doc.save(buf); buf.seek(0)
                docs_mem[fname_docx] = buf.read()

                pdfs_mem[fname_pdf] = build_pdf_document(row)

                log_event(run_id, "document_generated", investor, {"file": fname_docx})
                success += 1
                
                # Store individual docs for direct download
                doc_key = f"pe_doc_{run_id}_{safe_name}"
                pdf_key = f"pe_pdf_{run_id}_{safe_name}"
                _pe_results[doc_key] = docs_mem[fname_docx]
                _pe_results[pdf_key] = pdfs_mem[fname_pdf]

                event = {
                    "type": "progress", "investor": investor, "ok": True,
                    "verdict": verdict, "file": fname_docx, "error": None,
                    "doc_url": f"/api/pe/download/{doc_key}/file",
                    "pdf_url": f"/api/pe/download/{pdf_key}/file",
                }
            except Exception as exc:
                log_event(run_id, "document_failed", investor, {"reason": str(exc)})
                fail += 1
                event = {
                    "type": "progress", "investor": investor, "ok": False,
                    "verdict": verdict, "file": None, "error": str(exc),
                }

            yield f"data: {json.dumps(event)}\n\n"

        close_run(run_id, success, fail)

        # Build ZIPs + summary
        word_zip_buf = io.BytesIO()
        with zipfile.ZipFile(word_zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fn, d in docs_mem.items(): zf.writestr(fn, d)
        word_zip_buf.seek(0)

        pdf_zip_buf = io.BytesIO()
        with zipfile.ZipFile(pdf_zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fn, d in pdfs_mem.items(): zf.writestr(fn, d)
        pdf_zip_buf.seek(0)

        summary_df = build_summary_excel(df_raw, df_selected)
        excel_buf  = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl", date_format="YYYY-MM-DD") as w:
            summary_df.to_excel(w, index=False, sheet_name="CapitalStatements")
        excel_buf.seek(0)
        csv_bytes = summary_df.to_csv(index=False).encode("utf-8")

        _pe_results[run_id] = {
            "word_zip":      word_zip_buf.getvalue(),
            "pdf_zip":       pdf_zip_buf.getvalue(),
            "summary_excel": excel_buf.getvalue(),
            "summary_csv":   csv_bytes,
        }

        done_event = {
            "type":              "done",
            "run_id":            run_id,
            "success_count":     success,
            "fail_count":        fail,
            "word_zip_url":      f"/api/pe/download/{run_id}/word_zip",
            "pdf_zip_url":       f"/api/pe/download/{run_id}/pdf_zip",
            "summary_excel_url": f"/api/pe/download/{run_id}/summary_excel",
            "summary_csv_url":   f"/api/pe/download/{run_id}/summary_csv",
        }
        yield f"data: {json.dumps(done_event)}\n\n"

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/api/pe/download/{run_id}/{file_type}")
async def pe_download(run_id: str, file_type: str):
    """Serve a generated PE file by run_id or direct key."""
    # Check if run_id is actually a direct key (e.g. pe_doc_...)
    if run_id.startswith("pe_"):
        data = _pe_results.get(run_id)
        if not data:
            raise HTTPException(404, "File not found or expired")
        ext = "docx" if "_doc_" in run_id else "pdf"
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ext == "docx" else "application/pdf"
        fname = f"statement.{ext}"
        return Response(
            content=data,
            media_type=mime,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )

    store = _pe_results.get(run_id)
    if not store:
        raise HTTPException(404, "Run not found or expired")
    if file_type not in store:
        raise HTTPException(404, f"File type '{file_type}' not in this run")

    mime_map = {
        "word_zip":      ("application/zip",                                          "capital_statements_word.zip"),
        "pdf_zip":       ("application/zip",                                          "capital_statements_pdf.zip"),
        "summary_excel": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "capital_statements_summary.xlsx"),
        "summary_csv":   ("text/csv",                                                  "capital_statements_summary.csv"),
    }
    mime, fname = mime_map[file_type]
    return Response(
        content=store[file_type],
        media_type=mime,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


class PeInsightRequest(BaseModel):
    session_token: str
    scope: str            # "portfolio" | "investor"
    investor: str = ""
    insight_type: str = "full"
    session_data: list[dict] | None = None  # stateless fallback for serverless


@app.post("/api/pe/insights")
async def pe_insights(req: PeInsightRequest):
    sess = _pe_sessions.get(req.session_token)
    if sess:
        df_latest = sess["df_latest"]
    elif req.session_data:
        df_latest = _records_to_df(req.session_data)
    else:
        return {"ok": False, "error": "Session expired — please re-upload your file."}

    partnership = str(df_latest["PARTNERSHIP_NAME"].iloc[0]) if len(df_latest) else "—"
    as_of       = str(df_latest["TO_DATE"].max())[:10] if len(df_latest) else "—"

    if req.scope == "portfolio":
        prompt = _build_pe_portfolio_prompt(df_latest, partnership, as_of, req.insight_type)
        label  = f"{req.insight_type} — Portfolio"
    else:
        matches = df_latest[df_latest["INVESTOR_NAME"] == req.investor]
        if matches.empty:
            raise HTTPException(404, f"Investor '{req.investor}' not found in session")
        row    = matches.iloc[0]
        prompt = _build_pe_investor_prompt(row, req.insight_type)
        label  = f"{req.insight_type} — {req.investor}"

    try:
        t0 = time.time()
        markdown = await run_in_threadpool(_call_gemini_sync, prompt, _PE_SYSTEM)
        duration = round(time.time() - t0, 1)
        return {"ok": True, "label": label, "markdown": markdown, "duration_s": duration}
    except Exception as exc:
        return {"ok": False, "error": str(exc)}


class PeChatRequest(BaseModel):
    session_token: str
    messages: list[dict]


@app.post("/api/pe/chat")
async def pe_chat(req: PeChatRequest):
    sess = _pe_sessions.get(req.session_token)
    portfolio_ctx = ""
    if sess:
        df = sess["df_latest"]
        partnership = str(df["PARTNERSHIP_NAME"].iloc[0]) if len(df) else "—"
        portfolio_ctx = _portfolio_context(df, partnership)

    system = _PE_SYSTEM
    if portfolio_ctx:
        system += f"\n\nCURRENT PORTFOLIO CONTEXT:\n{portfolio_ctx}"

    reply = await run_in_threadpool(_call_gemini_chat_sync, req.messages, system)
    return {"ok": True, "reply": reply}


# ═══════════════════════════ HF ENDPOINTS ═════════════════════════════════════

@app.post("/api/hf/upload-pcap")
async def hf_upload_pcap(file: UploadFile = File(...)):
    content = await file.read()
    try:
        pcap_df = read_hf_pcap_from_upload(io.BytesIO(content))
    except Exception as exc:
        return {"ok": False, "error": str(exc)}

    pcap_token = str(uuid.uuid4())
    _hf_sessions[pcap_token] = {
        "pcap_df":  pcap_df,
        "filename": file.filename or "pcap.xlsx",
    }

    # Aggregate metrics
    def _col(col):
        if col in pcap_df.columns:
            v = pcap_df[col].dropna()
            return float(v.sum()) if len(v) else 0.0
        return 0.0

    def _avg(col):
        if col in pcap_df.columns:
            v = pcap_df[col].dropna()
            return float(v.mean()) if len(v) else None
        return None

    # Preview rows
    preview_cols = ["INVESTOR_NAME", "INCEPTION_DATE", "REPT_CCY",
                    "BEG_CAP_CQ", "END_CAP_CQ", "BEG_PX", "END_PX",
                    "GROSS_IRR", "NET_IRR", "TVPI", "LP_NET_WF",
                    "LOCKUP_MO", "LOCKUP_EXPIRED", "HWM_ACTIVE"]
    preview = []
    for _, row in pcap_df.iterrows():
        entry = {}
        for col in preview_cols:
            v = row.get(col)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                entry[col] = None
            elif hasattr(v, "item"):
                entry[col] = v.item()
            elif hasattr(v, "isoformat"):
                entry[col] = str(v)[:10]
            else:
                entry[col] = v
        preview.append(entry)

    return {
        "ok":            True,
        "pcap_token":    pcap_token,
        "n_investors":   len(pcap_df),
        "total_nav_cq":  _col("END_CAP_CQ"),
        "avg_net_irr":   _avg("NET_IRR"),
        "avg_tvpi":      _avg("TVPI"),
        "investors":     pcap_df["INVESTOR_NAME"].dropna().tolist(),
        "preview":       preview,
        "session_data":  _df_to_records(pcap_df),
        "error":         None,
    }


@app.post("/api/hf/upload-ledger")
async def hf_upload_ledger(file: UploadFile = File(...)):
    content = await file.read()
    try:
        cf_df = pd.read_excel(io.BytesIO(content))
    except Exception as exc:
        return {"ok": False, "error": str(exc)}

    required = {"INVESTOR_NAME", "TRANSACTION_DATE", "TYPE", "AMOUNT"}
    missing  = required - set(cf_df.columns)
    if missing:
        return {"ok": False, "error": f"Missing CF Ledger columns: {sorted(missing)}"}

    ledger_token = str(uuid.uuid4())
    _cf_sessions[ledger_token] = {"cf_df": cf_df}

    return {"ok": True, "ledger_token": ledger_token, "rows": len(cf_df), "error": None}


class HfGenerateRequest(BaseModel):
    pcap_token:   str
    ledger_token: str | None = None
    investors:    list[str]
    session_data: list[dict] | None = None  # stateless fallback for serverless


@app.post("/api/hf/generate")
async def hf_generate(req: HfGenerateRequest):
    """Build HF statements + companion Excel, stream results via SSE."""
    pcap_sess = _hf_sessions.get(req.pcap_token)
    if pcap_sess:
        pcap_df_all = pcap_sess["pcap_df"]
        filename    = pcap_sess["filename"]
    elif req.session_data:
        pcap_df_all = _records_to_df(req.session_data)
        filename    = "upload.xlsx"
    else:
        raise HTTPException(404, "PCAP session not found — re-upload PCAP file")

    cf_df = _cf_sessions.get(req.ledger_token or "", {}).get("cf_df")

    pcap_df = pcap_df_all[pcap_df_all["INVESTOR_NAME"].isin(req.investors)].copy()
    if pcap_df.empty:
        raise HTTPException(400, "No matching investors in PCAP session")

    async def event_stream() -> AsyncGenerator[str, None]:
        run_id = str(uuid.uuid4())
        hf_docs: dict[str, bytes] = {}
        hf_pdfs: dict[str, bytes] = {}
        success = fail = 0

        for _, row in pcap_df.iterrows():
            investor  = str(row["INVESTOR_NAME"]).strip()
            safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in investor)

            fname_docx = f"{safe_name}_hf_capital_statement.docx"
            fname_pdf  = f"{safe_name}_hf_capital_statement.pdf"

            try:
                def _build_hf(_row=row):
                    doc_obj = build_hf_docx(_row)
                    _buf = io.BytesIO(); doc_obj.save(_buf); _buf.seek(0)
                    return _buf.read(), build_hf_pdf(_row)

                word_bytes, pdf_bytes = await run_in_threadpool(_build_hf)
                hf_docs[fname_docx] = word_bytes
                hf_pdfs[fname_pdf]  = pdf_bytes
                
                # Store individual docs for direct download
                doc_key = f"hf_doc_{run_id}_{safe_name}"
                pdf_key = f"hf_pdf_{run_id}_{safe_name}"
                _hf_results[doc_key] = word_bytes
                _hf_results[pdf_key] = pdf_bytes

                success += 1
                event = {
                    "type": "progress", "investor": investor, "ok": True, "file": fname_docx,
                    "doc_url": f"/api/hf/download/{doc_key}/file",
                    "pdf_url": f"/api/hf/download/{pdf_key}/file",
                }
            except Exception as exc:
                fail += 1
                event = {"type": "progress", "investor": investor, "ok": False, "file": None, "error": str(exc)}

            yield f"data: {json.dumps(event)}\n\n"

        # Build companion Excel + ZIPs
        word_zip_buf = io.BytesIO()
        with zipfile.ZipFile(word_zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fn, d in hf_docs.items(): zf.writestr(fn, d)
        word_zip_buf.seek(0)

        pdf_zip_buf = io.BytesIO()
        with zipfile.ZipFile(pdf_zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fn, d in hf_pdfs.items(): zf.writestr(fn, d)
        pdf_zip_buf.seek(0)

        try:
            companion_xlsx = await run_in_threadpool(build_hf_workbook, pcap_df, cf_df)
        except Exception:
            companion_xlsx = None

        # Summary Excel
        summary_cols = [
            "INVESTOR_NAME", "INCEPTION_DATE", "REPT_CCY",
            "BEG_CAP_CQ", "END_CAP_CQ", "BEG_PX", "END_PX",
            "GROSS_IRR", "NET_IRR", "DPI", "RVPI", "TVPI",
            "HURDLE_AMT_ITD", "LP_NET_WF", "TOTAL_COMMIT",
        ]
        summ_df  = pcap_df[[c for c in summary_cols if c in pcap_df.columns]].copy()
        summ_buf = io.BytesIO()
        with pd.ExcelWriter(summ_buf, engine="openpyxl", date_format="YYYY-MM-DD") as w:
            summ_df.to_excel(w, index=False, sheet_name="HF_Summary")
        summ_buf.seek(0)

        _hf_results[run_id] = {
            "word_zip":       word_zip_buf.getvalue(),
            "pdf_zip":        pdf_zip_buf.getvalue(),
            "summary_excel":  summ_buf.getvalue(),
            "companion_xlsx": companion_xlsx,
        }

        done_event = {
            "type":               "done",
            "run_id":             run_id,
            "success_count":      success,
            "fail_count":         fail,
            "word_zip_url":       f"/api/hf/download/{run_id}/word_zip",
            "pdf_zip_url":        f"/api/hf/download/{run_id}/pdf_zip",
            "summary_excel_url":  f"/api/hf/download/{run_id}/summary_excel",
            "companion_xlsx_url": f"/api/hf/download/{run_id}/companion_xlsx" if companion_xlsx else None,
        }
        yield f"data: {json.dumps(done_event)}\n\n"

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/hf/download/{run_id}/{file_type}")
async def hf_download(run_id: str, file_type: str):
    """Serve a generated HF file by run_id or direct key."""
    # Check if run_id is actually a direct key (e.g. hf_doc_...)
    if run_id.startswith("hf_"):
        data = _hf_results.get(run_id)
        if not data:
            raise HTTPException(404, "File not found or expired")
        ext = "docx" if "_doc_" in run_id else "pdf"
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ext == "docx" else "application/pdf"
        fname = f"statement.{ext}"
        return Response(
            content=data,
            media_type=mime,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )

    store = _hf_results.get(run_id)
    if not store:
        raise HTTPException(404, "Run not found or expired")
    data = store.get(file_type)
    if data is None:
        raise HTTPException(404, f"'{file_type}' not available for this run")

    mime_map = {
        "word_zip":       ("application/zip",           "hf_statements_word.zip"),
        "pdf_zip":        ("application/zip",           "hf_statements_pdf.zip"),
        "summary_excel":  ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "hf_summary.xlsx"),
        "companion_xlsx": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "hf_pcap_model.xlsx"),
    }
    mime, fname = mime_map.get(file_type, ("application/octet-stream", file_type))
    return Response(
        content=data,
        media_type=mime,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


class HfInsightRequest(BaseModel):
    pcap_token:   str
    scope:        str
    investor:     str = ""
    insight_type: str = "full"
    session_data: list[dict] | None = None  # stateless fallback for serverless


@app.post("/api/hf/insights")
async def hf_insights(req: HfInsightRequest):
    sess = _hf_sessions.get(req.pcap_token)
    if sess:
        pcap_df = sess["pcap_df"]
    elif req.session_data:
        pcap_df = _records_to_df(req.session_data)
    else:
        return {"ok": False, "error": "Session expired — please re-upload your file."}

    if req.scope == "portfolio":
        prompt = _build_hf_portfolio_prompt(pcap_df, req.insight_type)
        label  = f"{req.insight_type} — Portfolio"
    else:
        matches = pcap_df[pcap_df["INVESTOR_NAME"] == req.investor]
        if matches.empty:
            raise HTTPException(404, f"Investor '{req.investor}' not found")
        prompt = _build_hf_investor_prompt(matches.iloc[0], req.insight_type)
        label  = f"{req.insight_type} — {req.investor}"

    try:
        t0 = time.time()
        markdown = await run_in_threadpool(_call_gemini_sync, prompt, _HF_SYSTEM)
        duration = round(time.time() - t0, 1)
        return {"ok": True, "label": label, "markdown": markdown, "duration_s": duration}
    except Exception as exc:
        return {"ok": False, "error": str(exc)}


class HfChatRequest(BaseModel):
    pcap_token: str
    messages:   list[dict]


@app.post("/api/hf/chat")
async def hf_chat(req: HfChatRequest):
    sess = _hf_sessions.get(req.pcap_token)
    system = _HF_SYSTEM
    if sess:
        pcap_df = sess["pcap_df"]
        total_nav = float(pcap_df["END_CAP_CQ"].dropna().sum()) if "END_CAP_CQ" in pcap_df.columns else 0
        avg_irr   = float(pcap_df["NET_IRR"].dropna().mean())    if "NET_IRR"    in pcap_df.columns else 0
        ctx = (f"Fund: {sess['filename']} · {len(pcap_df)} LPs · "
               f"Total NAV {fmt_usd(total_nav)} · Avg Net IRR {avg_irr:.1f}%")
        system += f"\n\nCURRENT FUND CONTEXT:\n{ctx}"

    reply = await run_in_threadpool(_call_gemini_chat_sync, req.messages, system)
    return {"ok": True, "reply": reply}


# ═══════════════════════════ AUDIT ════════════════════════════════════════════

@app.get("/api/audit/entries")
async def audit_entries():
    log_path = Path("/tmp/audit_log.jsonl") if Path("/tmp").exists() else Path("audit_log.jsonl")
    if not log_path.exists():
        return []
    entries = []
    for line in log_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rec = json.loads(line)
            # Normalise to the shape AuditLine expects
            entries.append({
                "ts":       rec.get("timestamp_utc", "")[:19].replace("T", " "),
                "type":     rec.get("event_type") or rec.get("status", "run"),
                "investor": rec.get("investor", rec.get("input_file", "")),
                "detail":   _audit_detail_str(rec),
            })
        except Exception:
            pass
    return entries


def _audit_detail_str(rec: dict) -> str:
    detail = rec.get("detail")
    if isinstance(detail, dict):
        return "  ".join(f"{k}: {v}" for k, v in detail.items())
    if detail:
        return str(detail)
    # fallback for run-level records
    parts = []
    if rec.get("investor_count") is not None:
        parts.append(f"investors: {rec['investor_count']}")
    if rec.get("input_file"):
        parts.append(f"file: {rec['input_file']}")
    if rec.get("success_count") is not None:
        parts.append(f"ok: {rec['success_count']}  fail: {rec['fail_count']}")
    return "  ".join(parts)


@app.get("/api/audit/download")
async def audit_download():
    log_path = Path("/tmp/audit_log.jsonl") if Path("/tmp").exists() else Path("audit_log.jsonl")
    content = log_path.read_bytes() if log_path.exists() else b""
    return Response(
        content=content,
        media_type="application/octet-stream",
        headers={"Content-Disposition": 'attachment; filename="audit_log.jsonl"'},
    )


# ── Health check ──────────────────────────────────────────────────────────────
@app.get("/api/health")
async def health():
    return {"ok": True, "gemini_key_set": bool(_GEMINI_KEY), "model": _GEMINI_MODEL}
