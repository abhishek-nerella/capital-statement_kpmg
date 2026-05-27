"""Hedge Fund Capital Account Statement — Word (.docx) and PDF builder.

Column names from hf_pcap_engine.read_hf_pcap():
  BEG_CAP_CQ/YTD/ITD, END_CAP_CQ/YTD/ITD, BEG_UNITS_*/END_UNITS_*,
  BEG_PX/END_PX, GROSS_IRR, NET_IRR, DPI/RVPI/TVPI,
  HURDLE_AMT_ITD, LP_NET_WF, LOCKUP_MO, CUSTODIAN, etc.
"""

from __future__ import annotations

import io
import pandas as pd

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from generate_capital_statements import fmt_usd, fmt_date, fmt_ratio


# ── Safe accessors ─────────────────────────────────────────────────────────────

def _g(row: pd.Series, field: str, default: float = 0.0) -> float:
    val = row.get(field, default)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _gs(row: pd.Series, field: str, default: str = "—") -> str:
    val = row.get(field, default)
    s = str(val).strip() if val is not None else ""
    return s if s and s not in ("nan", "None", "") else default


# ── Formatting helpers ─────────────────────────────────────────────────────────

def _fmt_pct(v) -> str:
    try:
        return f"{float(v):.2f}%"
    except Exception:
        return "—"


def _fmt_x(v) -> str:
    try:
        return f"{float(v):.2f}x"
    except Exception:
        return "—"


def _fmt_units(v) -> str:
    try:
        f = float(v)
        return f"{f:,.4f}" if f != 0.0 else "—"
    except Exception:
        return "—"


def _fmt_px(v) -> str:
    try:
        f = float(v)
        return f"${f:,.4f}" if f != 0.0 else "—"
    except Exception:
        return "—"


def _fmt_mo(v) -> str:
    try:
        mo = int(float(v))
        return f"{mo} months" if mo > 0 else "—"
    except Exception:
        return "—"


def _fmt_int(v, suffix: str = "") -> str:
    try:
        n = int(float(v))
        return f"{n}{suffix}" if n > 0 else "—"
    except Exception:
        return "—"


# ── KPMG brand constants ───────────────────────────────────────────────────────
_KPMG_BLUE     = RGBColor(0x00, 0x33, 0x8D)
_KPMG_BLUE_HEX = "#00338D"
_PACIFIC_HEX   = "#00B8F5"
_RED_HEX       = "#D73B3E"
_LIGHT_HEX     = "#ACEAFF"
_TOTAL_HEX     = "#E8F4FF"


# ── ReportLab styles (module-level, unique hf_ prefix) ────────────────────────
_hf_base = getSampleStyleSheet()
_hf_red  = rl_colors.HexColor(_RED_HEX)
_hf_blue = rl_colors.HexColor(_KPMG_BLUE_HEX)
_hf_pac  = rl_colors.HexColor(_PACIFIC_HEX)
_hf_lt   = rl_colors.HexColor(_LIGHT_HEX)

_HF_PS_CONF     = ParagraphStyle("hf_conf",     parent=_hf_base["Normal"], fontSize=9,  textColor=_hf_red,  alignment=TA_RIGHT, fontName="Helvetica-Bold")
_HF_PS_TITLE    = ParagraphStyle("hf_title",    parent=_hf_base["Normal"], fontSize=16, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=2)
_HF_PS_SUBTITLE = ParagraphStyle("hf_subtitle", parent=_hf_base["Normal"], fontSize=12, alignment=TA_CENTER, spaceAfter=4)
_HF_PS_NORMAL   = ParagraphStyle("hf_normal",   parent=_hf_base["Normal"], fontSize=10, leading=14)
_HF_PS_INDENT   = ParagraphStyle("hf_indent",   parent=_hf_base["Normal"], fontSize=10, leading=14, leftIndent=12)
_HF_PS_VALUE    = ParagraphStyle("hf_value",    parent=_hf_base["Normal"], fontSize=10, leading=14, alignment=TA_RIGHT)
_HF_PS_BOLD     = ParagraphStyle("hf_bold",     parent=_hf_base["Normal"], fontSize=10, leading=14, fontName="Helvetica-Bold")
_HF_PS_BOLDV    = ParagraphStyle("hf_boldv",    parent=_hf_base["Normal"], fontSize=10, leading=14, fontName="Helvetica-Bold", alignment=TA_RIGHT)
_HF_PS_FOOT     = ParagraphStyle("hf_foot",     parent=_hf_base["Normal"], fontSize=8,  fontName="Helvetica-Oblique")
_HF_PS_SEC      = ParagraphStyle("hf_sec",      parent=_hf_base["Normal"], fontSize=10, fontName="Helvetica-Bold", textColor=_hf_blue, spaceBefore=6, spaceAfter=3)
_HF_PS_TH       = ParagraphStyle("hf_th",       parent=_hf_base["Normal"], fontSize=8,  fontName="Helvetica-Bold", alignment=TA_CENTER, textColor=rl_colors.white)
_HF_PS_TH_L     = ParagraphStyle("hf_th_l",     parent=_hf_base["Normal"], fontSize=8,  fontName="Helvetica-Bold", textColor=rl_colors.white)
_HF_PS_TD_L     = ParagraphStyle("hf_td_l",     parent=_hf_base["Normal"], fontSize=9,  leading=12)
_HF_PS_TD_R     = ParagraphStyle("hf_td_r",     parent=_hf_base["Normal"], fontSize=9,  leading=12, alignment=TA_RIGHT)
_HF_PS_TD_LB    = ParagraphStyle("hf_td_lb",    parent=_hf_base["Normal"], fontSize=9,  leading=12, fontName="Helvetica-Bold")
_HF_PS_TD_RB    = ParagraphStyle("hf_td_rb",    parent=_hf_base["Normal"], fontSize=9,  leading=12, fontName="Helvetica-Bold", alignment=TA_RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# Word (.docx) helpers
# ══════════════════════════════════════════════════════════════════════════════

def _set_run(run, bold=False, underline=False, size_pt=10,
             color: RGBColor | None = None) -> None:
    run.bold = bold
    run.underline = underline
    run.font.size = Pt(size_pt)
    if color:
        run.font.color.rgb = color


def _two_col(doc: Document, label: str, value: str,
             bold: bool = False, indent: bool = False) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Inches(4.5), WD_ALIGN_PARAGRAPH.RIGHT)
    prefix = "    " if indent else ""
    run = p.add_run(f"{prefix}{label}\t{value}")
    _set_run(run, bold=bold)


def _section_hdr(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    _set_run(run, bold=True, underline=True, size_pt=10, color=_KPMG_BLUE)


def _kv_line(doc: Document, label: str, value: str) -> None:
    p = doc.add_paragraph()
    r1 = p.add_run(f"{label}: ")
    _set_run(r1, bold=True)
    r2 = p.add_run(value)
    _set_run(r2)


def _blank(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)


def _shaded_cell(cell, hex_fill: str) -> None:
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:fill"),  hex_fill)
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:val"),   "clear")
    tcPr.append(shd)


def _add_table_docx(doc: Document, headers: list[str],
                    data_rows: list[list[str]],
                    bold_last: bool = False) -> None:
    n_cols = len(headers)
    tbl = doc.add_table(rows=1 + len(data_rows), cols=n_cols)
    tbl.style = "Table Grid"

    hdr = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.bold           = True
            run.font.size      = Pt(8)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shaded_cell(hdr[i], "00338D")

    for ri, row_data in enumerate(data_rows):
        is_bold_row = bold_last and (ri == len(data_rows) - 1)
        cells = tbl.rows[ri + 1].cells
        for ci, val in enumerate(row_data):
            cells[ci].text = str(val)
            align = WD_ALIGN_PARAGRAPH.RIGHT if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT
            cells[ci].paragraphs[0].alignment = align
            for run in cells[ci].paragraphs[0].runs:
                run.font.size = Pt(9)
                run.bold      = is_bold_row
            if is_bold_row:
                _shaded_cell(cells[ci], "E8F4FF")


# ══════════════════════════════════════════════════════════════════════════════
# Word (.docx) builder
# ══════════════════════════════════════════════════════════════════════════════

def build_hf_docx(row: pd.Series) -> Document:
    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Inches(0.75)
        sec.bottom_margin = Inches(0.75)
        sec.left_margin   = Inches(1.0)
        sec.right_margin  = Inches(1.0)

    # ── CONFIDENTIAL + title ──────────────────────────────────────────────────
    p = doc.add_paragraph()
    r = p.add_run("CONFIDENTIAL")
    _set_run(r, bold=True, size_pt=9, color=RGBColor(0xD7, 0x3B, 0x3E))
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p = doc.add_paragraph()
    r = p.add_run("Hedge Fund Capital Account Statement")
    _set_run(r, bold=True, size_pt=16, color=_KPMG_BLUE)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    r = p.add_run("Current Reporting Period")
    _set_run(r, size_pt=11)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    _blank(doc)
    _kv_line(doc, "Investor",   _gs(row, "INVESTOR_NAME"))
    _kv_line(doc, "Inception",  _gs(row, "INCEPTION_DATE"))
    _kv_line(doc, "Currency",   _gs(row, "REPT_CCY"))
    _kv_line(doc, "% of AUM",   _fmt_pct(_g(row, "PCT_AUM")))
    _blank(doc)

    # ── 1: Capital Account Summary ────────────────────────────────────────────
    _section_hdr(doc, "1. Capital Account Summary")
    _blank(doc)
    _add_table_docx(doc,
        headers=["", "Current Quarter (CQ)", "Year-to-Date (YTD)", "Inception-to-Date (ITD)"],
        bold_last=True,
        data_rows=[
            ["Beginning Partner's Capital",    fmt_usd(_g(row,"BEG_CAP_CQ")),  fmt_usd(_g(row,"BEG_CAP_YTD")),  fmt_usd(_g(row,"BEG_CAP_ITD"))],
            ["(+) Capital Contributions",      fmt_usd(_g(row,"CONTRIB_CQ")),  fmt_usd(_g(row,"CONTRIB_YTD")),  fmt_usd(_g(row,"CONTRIB_ITD"))],
            ["(+) DRIP Reinvestment",          fmt_usd(_g(row,"DRIP_CQ")),     fmt_usd(_g(row,"DRIP_YTD")),     fmt_usd(_g(row,"DRIP_ITD"))],
            ["(+) Investment Income (Loss)",   fmt_usd(_g(row,"INC_CQ")),      fmt_usd(_g(row,"INC_YTD")),      fmt_usd(_g(row,"INC_ITD"))],
            ["(+) Net Unrealized Gain (Loss)", fmt_usd(_g(row,"UNRLZ_CQ")),    fmt_usd(_g(row,"UNRLZ_YTD")),    fmt_usd(_g(row,"UNRLZ_ITD"))],
            ["(+) Net Realized Gain (Loss)",   fmt_usd(_g(row,"RLZD_CQ")),     fmt_usd(_g(row,"RLZD_YTD")),     fmt_usd(_g(row,"RLZD_ITD"))],
            ["(–) Distributions to LP",  fmt_usd(_g(row,"DIST_LP_CQ")),  fmt_usd(_g(row,"DIST_LP_YTD")),  fmt_usd(_g(row,"DIST_LP_ITD"))],
            ["(–) Incentive Fees",        fmt_usd(_g(row,"INC_FEE_CQ")),  fmt_usd(_g(row,"INC_FEE_YTD")),  fmt_usd(_g(row,"INC_FEE_ITD"))],
            ["Ending Partner's Capital",       fmt_usd(_g(row,"END_CAP_CQ")),  fmt_usd(_g(row,"END_CAP_YTD")),  fmt_usd(_g(row,"END_CAP_ITD"))],
        ],
    )
    _blank(doc)

    # ── 2: Unit Reconciliation ────────────────────────────────────────────────
    _section_hdr(doc, "2. Unit Reconciliation")
    _blank(doc)
    _add_table_docx(doc,
        headers=["", "CQ Units", "YTD Units", "ITD Units"],
        bold_last=True,
        data_rows=[
            ["Beginning Units",    _fmt_units(_g(row,"BEG_UNITS_CQ")),  _fmt_units(_g(row,"BEG_UNITS_YTD")),  _fmt_units(_g(row,"BEG_UNITS_ITD"))],
            ["(+) Units Issued",   _fmt_units(_g(row,"CONTRIB_U_CQ")),  _fmt_units(_g(row,"CONTRIB_U_YTD")),  _fmt_units(_g(row,"CONTRIB_U_ITD"))],
            ["(–) Units Redeemed", _fmt_units(_g(row,"REDEMP_U_CQ")),   _fmt_units(_g(row,"REDEMP_U_YTD")),   _fmt_units(_g(row,"REDEMP_U_ITD"))],
            ["Ending Units",       _fmt_units(_g(row,"END_UNITS_CQ")),  _fmt_units(_g(row,"END_UNITS_YTD")),  _fmt_units(_g(row,"END_UNITS_ITD"))],
        ],
    )
    _blank(doc)

    # ── 3: Unit Price Attribution ─────────────────────────────────────────────
    _section_hdr(doc, "3. Unit Price Attribution")
    _blank(doc)
    _add_table_docx(doc,
        headers=["Line Item", "Unit Price"],
        bold_last=True,
        data_rows=[
            ["Beginning Partner's Capital",            _fmt_px(_g(row,"BEG_PX"))],
            ["Transfer In",                            _fmt_px(_g(row,"XFER_IN_PX"))],
            ["Transfer Out",                           _fmt_px(_g(row,"XFER_OUT_PX"))],
            ["Capital Contribution",                   _fmt_px(_g(row,"CONTRIB_PX"))],
            ["Investment Income",                      _fmt_px(_g(row,"INC_PX"))],
            ["Fund Level Expense",                     _fmt_px(_g(row,"EXP_PX"))],
            ["Net Unrealized Gain / (Loss)",           _fmt_px(_g(row,"UNRLZ_PX"))],
            ["Net Realized Gain / (Loss)",             _fmt_px(_g(row,"RLZD_PX"))],
            ["Partner's Equity Before Distributions",  _fmt_px(_g(row,"EQ_PRED_PX"))],
            ["Distributions to LPs",                  _fmt_px(_g(row,"DIST_LP_PX"))],
            ["Distributions to Manager",               _fmt_px(_g(row,"DIST_MGR_PX"))],
            ["Incentive Fee",                          _fmt_px(_g(row,"INC_FEE_PX"))],
            ["Tax Reduction on Distributions",         _fmt_px(_g(row,"TAX_RED_PX"))],
            ["Ending Partner's Capital",               _fmt_px(_g(row,"END_PX"))],
        ],
    )
    _blank(doc)

    # ── 4: Performance Analytics ──────────────────────────────────────────────
    _section_hdr(doc, "4. Performance Analytics")
    _blank(doc)
    _two_col(doc, "Gross IRR",           _fmt_pct(_g(row,"GROSS_IRR")))
    _two_col(doc, "Net IRR",             _fmt_pct(_g(row,"NET_IRR")))
    _two_col(doc, "DPI",                 _fmt_x(_g(row,"DPI")))
    _two_col(doc, "RVPI",                _fmt_x(_g(row,"RVPI")))
    _two_col(doc, "TVPI",                _fmt_x(_g(row,"TVPI")))
    _two_col(doc, "Total Return (CQ)",   _fmt_pct(_g(row,"TOT_RET_CQ_PCT")))
    _two_col(doc, "Total Return (ITD)",  _fmt_pct(_g(row,"TOT_RET_ITD_PCT")))
    _two_col(doc, "Preferred Return",    _fmt_pct(_g(row,"PREF_RET")))
    _two_col(doc, "Incentive Fee Rate",  _fmt_pct(_g(row,"INC_FEE_RATE")))
    _two_col(doc, "Hurdle Exceeded?",    _gs(row,"HURDLE_EXCEEDED"))
    _blank(doc)

    # ── 5: Capital Commitment ─────────────────────────────────────────────────
    _section_hdr(doc, "5. Capital Commitment")
    _blank(doc)
    _two_col(doc, "Total Commitment",     fmt_usd(_g(row,"TOTAL_COMMIT")))
    _two_col(doc, "Funded Commitment",    fmt_usd(_g(row,"FUNDED_COMMIT")))
    _two_col(doc, "Transferred",          fmt_usd(_g(row,"XFER_COMMIT")))
    _two_col(doc, "Available Commitment", fmt_usd(_g(row,"AVAIL_COMMIT")))
    _two_col(doc, "Commitment Funded",    _fmt_pct(_g(row,"COMMIT_FUNDED_PCT")))
    _blank(doc)

    # ── 6: Return & Fee Analysis ──────────────────────────────────────────────
    _section_hdr(doc, "6. Return & Fee Analysis")
    _blank(doc)
    _add_table_docx(doc,
        headers=["Metric", "CQ ($)", "ITD ($)"],
        data_rows=[
            ["Unrealized Gain",  fmt_usd(_g(row,"UNRLZ_CQ_DLR")),  fmt_usd(_g(row,"UNRLZ_ITD_DLR"))],
            ["Realized Gain",    fmt_usd(_g(row,"RLZD_CQ_DLR")),   fmt_usd(_g(row,"RLZD_ITD_DLR"))],
            ["Total Return",     fmt_usd(_g(row,"TOT_RET_CQ_DLR")), fmt_usd(_g(row,"TOT_RET_ITD_DLR"))],
            ["Mgmt. Fee",        "—",                               fmt_usd(_g(row,"MGMT_FEE_ITD_DLR"))],
            ["Incentive Fee",    "—",                               fmt_usd(_g(row,"INC_FEE_ITD_DLR"))],
            ["Total Fees",       "—",                               fmt_usd(_g(row,"TOT_FEES_ITD_DLR"))],
            ["Net Return",       "—",                               fmt_usd(_g(row,"NET_RET_ITD_DLR"))],
        ],
        bold_last=True,
    )
    _blank(doc)

    # ── 7: Waterfall Analysis ─────────────────────────────────────────────────
    _section_hdr(doc, "7. Waterfall Analysis")
    _blank(doc)
    _two_col(doc, "Hurdle Amount (ITD)",     fmt_usd(_g(row,"HURDLE_AMT_ITD")))
    _two_col(doc, "Excess Over Hurdle",      fmt_usd(_g(row,"EXCESS_HURDLE")))
    _two_col(doc, "GP Catch-Up Amount",      fmt_usd(_g(row,"GP_CATCHUP_AMT")))
    _two_col(doc, "LP Net Waterfall Share",  fmt_usd(_g(row,"LP_NET_WF")), bold=True)
    _two_col(doc, "Waterfall Tier",          _gs(row,"WF_TIER"))
    _blank(doc)

    # ── 8: Lock-up & Redemption Terms ────────────────────────────────────────
    _section_hdr(doc, "8. Lock-up & Redemption Terms")
    _blank(doc)
    _two_col(doc, "Lock-up Period",               _fmt_mo(_g(row,"LOCKUP_MO")))
    _two_col(doc, "Subscription Date",            _gs(row,"SUB_DATE"))
    _two_col(doc, "First Contribution Date",      _gs(row,"FIRST_CONTRIB_DATE"))
    _two_col(doc, "Redemption Eligibility Date",  _gs(row,"REDEMP_ELIG_DATE"))
    _two_col(doc, "Lock-up Expired?",             _gs(row,"LOCKUP_EXPIRED"))
    _two_col(doc, "Months Remaining",             _fmt_int(_g(row,"MONTHS_REM"), " mo"))
    _two_col(doc, "Redemption Frequency",         _gs(row,"REDEMP_FREQ"))
    _two_col(doc, "Gate Provision",               _gs(row,"GATE_PROV"))
    _two_col(doc, "Notice Period",                _fmt_int(_g(row,"NOTICE_DAYS"), " days"))
    _two_col(doc, "Side Pocket Eligible?",        _gs(row,"SIDE_POCKET"))
    _two_col(doc, "DRIP Enrolled?",               _gs(row,"DRIP_ENROLLED"))
    _two_col(doc, "Distribution Preference",      _gs(row,"DIST_PREF"))
    _two_col(doc, "High-Water Mark Active?",      _gs(row,"HWM_ACTIVE"))
    _blank(doc)

    # ── 9: Side Letter & Compliance ───────────────────────────────────────────
    _section_hdr(doc, "9. Side Letter & Compliance")
    _blank(doc)
    _two_col(doc, "Side Letter Flag",               _gs(row,"SIDE_LETTER_FLG"))
    _two_col(doc, "Mgmt. Fee Rate (Side Letter)",   _fmt_pct(_g(row,"MGMT_FEE_RATE")))
    _two_col(doc, "Hurdle Type",                    _gs(row,"HURDLE_TYPE"))
    _two_col(doc, "Catch-up %",                     _fmt_pct(_g(row,"CATCHUP_PCT")))
    _two_col(doc, "Reporting Currency",             _gs(row,"REPT_CCY"))
    _two_col(doc, "FATCA Status",                   _gs(row,"FATCA"))
    _two_col(doc, "AML / KYC Status",               _gs(row,"AML_KYC"))
    _two_col(doc, "Accredited / Qualified",         _gs(row,"ACCREDITED"))
    _two_col(doc, "Custodian / Prime Broker",       _gs(row,"CUSTODIAN"))
    _two_col(doc, "Special Terms",                  _gs(row,"SPECIAL_TERMS"))
    _blank(doc)

    # ── Footer ────────────────────────────────────────────────────────────────
    p = doc.add_paragraph(
        "This document is CONFIDENTIAL and prepared for the named investor only. "
        "Capital account balances are based on pre-calculated PCAP data and may not "
        "represent amounts ultimately realized. © KPMG International"
    )
    for run in p.runs:
        run.font.size = Pt(8)

    return doc


# ══════════════════════════════════════════════════════════════════════════════
# PDF builder
# ══════════════════════════════════════════════════════════════════════════════

def build_hf_pdf(row: pd.Series) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.75*inch, bottomMargin=0.75*inch,
        leftMargin=1.0*inch, rightMargin=1.0*inch,
    )

    KB  = rl_colors.HexColor(_KPMG_BLUE_HEX)
    TOT = rl_colors.HexColor(_TOTAL_HEX)
    W   = 6.5 * inch   # usable page width

    # ── PDF table builders ────────────────────────────────────────────────────

    def _kv(label, value):
        return Paragraph(f"<b>{label}:</b> {value}", _HF_PS_NORMAL)

    def _sec(text):
        return Paragraph(f"<b><u>{text}</u></b>", _HF_PS_SEC)

    def _two(label, value, bold=False):
        ls = _HF_PS_BOLD  if bold else _HF_PS_NORMAL
        vs = _HF_PS_BOLDV if bold else _HF_PS_VALUE
        t = Table([[Paragraph(label, ls), Paragraph(value, vs)]],
                  colWidths=[W * 0.62, W * 0.38])
        t.setStyle(TableStyle([
            ("TOPPADDING",    (0,0), (-1,-1), 1),
            ("BOTTOMPADDING", (0,0), (-1,-1), 1),
            ("LEFTPADDING",   (0,0), (-1,-1), 0),
            ("RIGHTPADDING",  (-1,0),(-1,-1), 0),
        ]))
        return t

    def _grid(headers, rows, col_widths, bold_last=False):
        data = [[Paragraph(h, _HF_PS_TH) for h in headers]]
        for ri, rd in enumerate(rows):
            is_bold = bold_last and (ri == len(rows) - 1)
            sl = _HF_PS_TD_LB if is_bold else _HF_PS_TD_L
            sr = _HF_PS_TD_RB if is_bold else _HF_PS_TD_R
            data.append([
                Paragraph(str(v), sr if i > 0 else sl)
                for i, v in enumerate(rd)
            ])
        t = Table(data, colWidths=col_widths)
        ts = [
            ("BACKGROUND",    (0, 0), (-1, 0), KB),
            ("TEXTCOLOR",     (0, 0), (-1, 0), rl_colors.white),
            ("GRID",          (0, 0), (-1, -1), 0.4, _hf_lt),
            ("BACKGROUND",    (0, 1), (-1, -1), rl_colors.white),
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING",   (0, 0), (-1, -1), 4),
            ("RIGHTPADDING",  (-1,0), (-1, -1), 4),
        ]
        if bold_last and len(rows) > 0:
            lr = len(rows)
            ts.append(("BACKGROUND", (0, lr), (-1, lr), TOT))
        t.setStyle(TableStyle(ts))
        return t

    sp = lambda h=5: Spacer(1, h)
    story = []

    # Header block
    story += [Paragraph("CONFIDENTIAL", _HF_PS_CONF), sp(4)]
    story += [Paragraph("Hedge Fund Capital Account Statement", _HF_PS_TITLE), sp(4)]
    story += [Paragraph("Current Reporting Period", _HF_PS_SUBTITLE), sp(8)]
    story += [_kv("Investor",  _gs(row,"INVESTOR_NAME")), sp(3)]
    story += [_kv("Inception", _gs(row,"INCEPTION_DATE")), sp(3)]
    story += [_kv("Currency",  _gs(row,"REPT_CCY")), sp(3)]
    story += [_kv("% of AUM",  _fmt_pct(_g(row,"PCT_AUM"))), sp(10)]

    # Section 1
    CW4 = [W*0.34, W*0.22, W*0.22, W*0.22]
    story += [_sec("1. Capital Account Summary"), sp(4)]
    story.append(_grid(
        ["", "Current Quarter", "Year-to-Date", "Inception-to-Date"],
        [
            ["Beginning Partner's Capital",    fmt_usd(_g(row,"BEG_CAP_CQ")),  fmt_usd(_g(row,"BEG_CAP_YTD")),  fmt_usd(_g(row,"BEG_CAP_ITD"))],
            ["(+) Capital Contributions",      fmt_usd(_g(row,"CONTRIB_CQ")),  fmt_usd(_g(row,"CONTRIB_YTD")),  fmt_usd(_g(row,"CONTRIB_ITD"))],
            ["(+) DRIP Reinvestment",          fmt_usd(_g(row,"DRIP_CQ")),     fmt_usd(_g(row,"DRIP_YTD")),     fmt_usd(_g(row,"DRIP_ITD"))],
            ["(+) Investment Income (Loss)",   fmt_usd(_g(row,"INC_CQ")),      fmt_usd(_g(row,"INC_YTD")),      fmt_usd(_g(row,"INC_ITD"))],
            ["(+) Net Unrealized Gain (Loss)", fmt_usd(_g(row,"UNRLZ_CQ")),    fmt_usd(_g(row,"UNRLZ_YTD")),    fmt_usd(_g(row,"UNRLZ_ITD"))],
            ["(+) Net Realized Gain (Loss)",   fmt_usd(_g(row,"RLZD_CQ")),     fmt_usd(_g(row,"RLZD_YTD")),     fmt_usd(_g(row,"RLZD_ITD"))],
            ["(–) Distributions to LP",  fmt_usd(_g(row,"DIST_LP_CQ")),  fmt_usd(_g(row,"DIST_LP_YTD")),  fmt_usd(_g(row,"DIST_LP_ITD"))],
            ["(–) Incentive Fees",        fmt_usd(_g(row,"INC_FEE_CQ")),  fmt_usd(_g(row,"INC_FEE_YTD")),  fmt_usd(_g(row,"INC_FEE_ITD"))],
            ["Ending Partner's Capital",       fmt_usd(_g(row,"END_CAP_CQ")),  fmt_usd(_g(row,"END_CAP_YTD")),  fmt_usd(_g(row,"END_CAP_ITD"))],
        ],
        col_widths=CW4, bold_last=True,
    ))
    story.append(sp(10))

    # Section 2
    story += [_sec("2. Unit Reconciliation"), sp(4)]
    story.append(_grid(
        ["", "CQ Units", "YTD Units", "ITD Units"],
        [
            ["Beginning Units",    _fmt_units(_g(row,"BEG_UNITS_CQ")),  _fmt_units(_g(row,"BEG_UNITS_YTD")),  _fmt_units(_g(row,"BEG_UNITS_ITD"))],
            ["(+) Units Issued",   _fmt_units(_g(row,"CONTRIB_U_CQ")),  _fmt_units(_g(row,"CONTRIB_U_YTD")),  _fmt_units(_g(row,"CONTRIB_U_ITD"))],
            ["(–) Units Redeemed", _fmt_units(_g(row,"REDEMP_U_CQ")),   _fmt_units(_g(row,"REDEMP_U_YTD")),   _fmt_units(_g(row,"REDEMP_U_ITD"))],
            ["Ending Units",       _fmt_units(_g(row,"END_UNITS_CQ")),  _fmt_units(_g(row,"END_UNITS_YTD")),  _fmt_units(_g(row,"END_UNITS_ITD"))],
        ],
        col_widths=CW4, bold_last=True,
    ))
    story.append(sp(10))

    # Section 3
    CW2A = [W * 0.55, W * 0.45]
    story += [_sec("3. Unit Price Attribution"), sp(4)]
    story.append(_grid(
        ["Line Item", "Unit Price"],
        [
            ["Beginning Partner's Capital",           _fmt_px(_g(row,"BEG_PX"))],
            ["Transfer In",                           _fmt_px(_g(row,"XFER_IN_PX"))],
            ["Transfer Out",                          _fmt_px(_g(row,"XFER_OUT_PX"))],
            ["Capital Contribution",                  _fmt_px(_g(row,"CONTRIB_PX"))],
            ["Investment Income",                     _fmt_px(_g(row,"INC_PX"))],
            ["Fund Level Expense",                    _fmt_px(_g(row,"EXP_PX"))],
            ["Net Unrealized Gain / (Loss)",          _fmt_px(_g(row,"UNRLZ_PX"))],
            ["Net Realized Gain / (Loss)",            _fmt_px(_g(row,"RLZD_PX"))],
            ["Partner's Equity Before Distributions", _fmt_px(_g(row,"EQ_PRED_PX"))],
            ["Distributions to LPs",                 _fmt_px(_g(row,"DIST_LP_PX"))],
            ["Distributions to Manager",              _fmt_px(_g(row,"DIST_MGR_PX"))],
            ["Incentive Fee",                         _fmt_px(_g(row,"INC_FEE_PX"))],
            ["Tax Reduction on Distributions",        _fmt_px(_g(row,"TAX_RED_PX"))],
            ["Ending Partner's Capital",              _fmt_px(_g(row,"END_PX"))],
        ],
        col_widths=CW2A, bold_last=True,
    ))
    story.append(sp(10))

    # Section 4
    story += [_sec("4. Performance Analytics"), sp(4)]
    for lbl, val in [
        ("Gross IRR",           _fmt_pct(_g(row,"GROSS_IRR"))),
        ("Net IRR",             _fmt_pct(_g(row,"NET_IRR"))),
        ("DPI",                 _fmt_x(_g(row,"DPI"))),
        ("RVPI",                _fmt_x(_g(row,"RVPI"))),
        ("TVPI",                _fmt_x(_g(row,"TVPI"))),
        ("Total Return (CQ)",   _fmt_pct(_g(row,"TOT_RET_CQ_PCT"))),
        ("Total Return (ITD)",  _fmt_pct(_g(row,"TOT_RET_ITD_PCT"))),
        ("Preferred Return",    _fmt_pct(_g(row,"PREF_RET"))),
        ("Incentive Fee Rate",  _fmt_pct(_g(row,"INC_FEE_RATE"))),
        ("Hurdle Exceeded?",    _gs(row,"HURDLE_EXCEEDED")),
    ]:
        story.append(_two(lbl, val))
    story.append(sp(10))

    # Section 5
    story += [_sec("5. Capital Commitment"), sp(4)]
    for lbl, val in [
        ("Total Commitment",     fmt_usd(_g(row,"TOTAL_COMMIT"))),
        ("Funded Commitment",    fmt_usd(_g(row,"FUNDED_COMMIT"))),
        ("Transferred",          fmt_usd(_g(row,"XFER_COMMIT"))),
        ("Available Commitment", fmt_usd(_g(row,"AVAIL_COMMIT"))),
        ("Commitment Funded",    _fmt_pct(_g(row,"COMMIT_FUNDED_PCT"))),
    ]:
        story.append(_two(lbl, val))
    story.append(sp(10))

    # Section 6
    CW3 = [W*0.40, W*0.30, W*0.30]
    story += [_sec("6. Return & Fee Analysis"), sp(4)]
    story.append(_grid(
        ["Metric", "CQ ($)", "ITD ($)"],
        [
            ["Unrealized Gain",  fmt_usd(_g(row,"UNRLZ_CQ_DLR")),   fmt_usd(_g(row,"UNRLZ_ITD_DLR"))],
            ["Realized Gain",    fmt_usd(_g(row,"RLZD_CQ_DLR")),    fmt_usd(_g(row,"RLZD_ITD_DLR"))],
            ["Total Return",     fmt_usd(_g(row,"TOT_RET_CQ_DLR")), fmt_usd(_g(row,"TOT_RET_ITD_DLR"))],
            ["Mgmt. Fee",        "—",                                fmt_usd(_g(row,"MGMT_FEE_ITD_DLR"))],
            ["Incentive Fee",    "—",                                fmt_usd(_g(row,"INC_FEE_ITD_DLR"))],
            ["Total Fees",       "—",                                fmt_usd(_g(row,"TOT_FEES_ITD_DLR"))],
            ["Net Return",       "—",                                fmt_usd(_g(row,"NET_RET_ITD_DLR"))],
        ],
        col_widths=CW3, bold_last=True,
    ))
    story.append(sp(10))

    # Section 7
    story += [_sec("7. Waterfall Analysis"), sp(4)]
    for lbl, val, bold in [
        ("Hurdle Amount (ITD)",     fmt_usd(_g(row,"HURDLE_AMT_ITD")), False),
        ("Excess Over Hurdle",      fmt_usd(_g(row,"EXCESS_HURDLE")),  False),
        ("GP Catch-Up Amount",      fmt_usd(_g(row,"GP_CATCHUP_AMT")), False),
        ("LP Net Waterfall Share",  fmt_usd(_g(row,"LP_NET_WF")),      True),
        ("Waterfall Tier",          _gs(row,"WF_TIER"),                False),
    ]:
        story.append(_two(lbl, val, bold=bold))
    story.append(sp(10))

    # Section 8
    story += [_sec("8. Lock-up & Redemption Terms"), sp(4)]
    for lbl, val in [
        ("Lock-up Period",              _fmt_mo(_g(row,"LOCKUP_MO"))),
        ("Subscription Date",           _gs(row,"SUB_DATE")),
        ("First Contribution Date",     _gs(row,"FIRST_CONTRIB_DATE")),
        ("Redemption Eligibility Date", _gs(row,"REDEMP_ELIG_DATE")),
        ("Lock-up Expired?",            _gs(row,"LOCKUP_EXPIRED")),
        ("Months Remaining",            _fmt_int(_g(row,"MONTHS_REM"), " mo")),
        ("Redemption Frequency",        _gs(row,"REDEMP_FREQ")),
        ("Gate Provision",              _gs(row,"GATE_PROV")),
        ("Notice Period",               _fmt_int(_g(row,"NOTICE_DAYS"), " days")),
        ("Side Pocket Eligible?",       _gs(row,"SIDE_POCKET")),
        ("DRIP Enrolled?",              _gs(row,"DRIP_ENROLLED")),
        ("Distribution Preference",     _gs(row,"DIST_PREF")),
        ("High-Water Mark Active?",     _gs(row,"HWM_ACTIVE")),
    ]:
        story.append(_two(lbl, val))
    story.append(sp(10))

    # Section 9
    story += [_sec("9. Side Letter & Compliance"), sp(4)]
    for lbl, val in [
        ("Side Letter Flag",              _gs(row,"SIDE_LETTER_FLG")),
        ("Mgmt. Fee Rate (Side Letter)",  _fmt_pct(_g(row,"MGMT_FEE_RATE"))),
        ("Hurdle Type",                   _gs(row,"HURDLE_TYPE")),
        ("Catch-up %",                    _fmt_pct(_g(row,"CATCHUP_PCT"))),
        ("Reporting Currency",            _gs(row,"REPT_CCY")),
        ("FATCA Status",                  _gs(row,"FATCA")),
        ("AML / KYC Status",              _gs(row,"AML_KYC")),
        ("Accredited / Qualified",        _gs(row,"ACCREDITED")),
        ("Custodian / Prime Broker",      _gs(row,"CUSTODIAN")),
        ("Special Terms",                 _gs(row,"SPECIAL_TERMS")),
    ]:
        story.append(_two(lbl, val))
    story.append(sp(16))

    story.append(Paragraph(
        "This statement is CONFIDENTIAL and prepared for the named investor only. "
        "Balances are based on pre-calculated PCAP data and may not represent amounts "
        "ultimately realized. © KPMG International",
        _HF_PS_FOOT,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()
