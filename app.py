"""
KPMG Capital Statement Generator — Streamlit UI
Run: streamlit run app.py
"""

import io
import os
import zipfile

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
    raise EnvironmentError("GEMINI_API_KEY environment variable is not set.")
_GEMINI_MODEL = "models/gemini-2.5-flash"

from generate_capital_statements import (
    build_document, build_summary_excel,
    fmt_usd, fmt_ratio, fmt_date,
    REQUIRED_COLS as _REQUIRED_COLS,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Capital Analysis Statement Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Brand tokens ──────────────────────────────────────────────────────────────
NAVY   = "#00338D"
SKY    = "#0091DA"
LIGHT  = "#E8F4FD"
BORDER = "#D0D9E8"
TEXT   = "#1A1A2E"
MUTED  = "#6B7A99"
WHITE  = "#FFFFFF"
RED    = "#D73B3E"
GOLD   = "#B5962D"

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


# ── PDF builder ───────────────────────────────────────────────────────────────

def build_pdf_document(row: pd.Series) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.75*inch, bottomMargin=0.75*inch,
        leftMargin=1.0*inch, rightMargin=1.0*inch,
    )
    COL_W = [4.25*inch, 2.25*inch]

    def _two_col(label: str, value: str, indented: bool = False) -> Table:
        lbl_s = _PS_INDENT if indented else _PS_NORMAL
        data  = [[Paragraph(label, lbl_s), Paragraph(value, _PS_VALUE)]]
        t = Table(data, colWidths=COL_W)
        t.setStyle(TableStyle([
            ("TOPPADDING",    (0,0), (-1,-1), 1),
            ("BOTTOMPADDING", (0,0), (-1,-1), 1),
            ("LEFTPADDING",   (0,0), (-1,-1), 0),
            ("RIGHTPADDING",  (-1,0), (-1,-1), 0),
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
    story.append(_two_col("Capital contributed to date (B)",                  fmt_usd(contributed)))
    story.append(_two_col("Remaining capital commitment (A-B)",               fmt_usd(committed - contributed)))
    story.append(sp(16))

    story += [_section("Summary of Distributions and Valuation"), sp(6)]
    story.append(_two_col("Total capital contributed to date",                                   fmt_usd(row["INCEPTION_TO_DATE_CONTRIBUTION"])))
    story.append(_two_col("Total distributions to date",                                          fmt_usd(row["INCEPTION_TO_DATE_DISTRIBUTION"])))
    story.append(_two_col(f"Capital balance (remaining value) at {fmt_date(row['TO_DATE'])} *",  fmt_usd(row["CLOSING_YTD_NAV"])))
    story.append(_two_col("Total Estimated Value (distributions + balance)",                      fmt_usd(row["TEV"])))
    story.append(_two_col("Total Estimated Value as net multiple",                                fmt_ratio(row["TEV_RATIO"])))
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


def _call_gemini(prompt: str) -> str:
    client = _genai.Client(api_key=_GEMINI_KEY)
    response = client.models.generate_content(
        model=_GEMINI_MODEL,
        config=_genai.types.GenerateContentConfig(
            system_instruction=_PE_SYSTEM,
            max_output_tokens=1800,
        ),
        contents=prompt,
    )
    return response.text


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
            max_output_tokens=800,
        ),
        contents=contents,
    )
    return response.text


# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; color: {TEXT}; }}
  .stApp {{ background: #F4F6FA; }}

  /* Sidebar */
  section[data-testid="stSidebar"] {{
    background: {NAVY} !important;
    border-right: none;
  }}
  section[data-testid="stSidebar"] * {{ color: {WHITE} !important; }}
  section[data-testid="stSidebar"] .stFileUploader {{
    background: rgba(255,255,255,0.08);
    border: 1px dashed rgba(255,255,255,0.35);
    border-radius: 6px; padding: 8px;
  }}
  section[data-testid="stSidebar"] .stTextInput input,
  section[data-testid="stSidebar"] textarea {{
    background: rgba(255,255,255,0.12) !important;
    border: 1px solid rgba(255,255,255,0.30) !important;
    color: {WHITE} !important; border-radius: 4px;
    font-size: 13px !important;
  }}
  section[data-testid="stSidebar"] .stTextArea textarea {{
    background: rgba(255,255,255,0.10) !important;
    border: 1px solid rgba(255,255,255,0.25) !important;
    color: {WHITE} !important; border-radius: 4px;
    font-size: 13px !important; resize: none;
  }}

  .sidebar-divider {{
    border: none; border-top: 1px solid rgba(255,255,255,0.18); margin: 14px 0;
  }}

  /* Chat messages in sidebar */
  .chat-bubble-user {{
    background: rgba(0,145,218,0.25);
    border-radius: 10px 10px 2px 10px;
    padding: 7px 11px; margin: 4px 0 4px 20px;
    font-size: 12px; color: {WHITE}; word-wrap: break-word;
  }}
  .chat-bubble-ai {{
    background: rgba(255,255,255,0.12);
    border-radius: 10px 10px 10px 2px;
    padding: 7px 11px; margin: 4px 20px 4px 0;
    font-size: 12px; color: rgba(255,255,255,0.88); word-wrap: break-word;
  }}
  .chat-label-user {{ font-size: 10px; opacity: 0.55; text-align: right; margin-bottom: 1px; }}
  .chat-label-ai   {{ font-size: 10px; opacity: 0.55; margin-bottom: 1px; }}

  /* Metric cards */
  .metric-card {{
    background: {WHITE}; border: 1px solid {BORDER};
    border-top: 3px solid {NAVY}; border-radius: 6px;
    padding: 18px 20px; margin-bottom: 12px;
  }}
  .metric-card .label {{
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.08em; color: {MUTED}; margin-bottom: 4px;
  }}
  .metric-card .value {{ font-size: 28px; font-weight: 700; color: {NAVY}; line-height: 1.1; }}
  .metric-card .sub   {{ font-size: 12px; color: {MUTED}; margin-top: 2px; }}

  .section-header {{
    font-size: 13px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.10em; color: {MUTED}; margin: 28px 0 12px 0;
    display: flex; align-items: center; gap: 8px;
  }}
  .section-header::after {{ content: ''; flex: 1; height: 1px; background: {BORDER}; }}

  .investor-pill {{
    display: inline-block; background: {LIGHT}; border: 1px solid {SKY};
    color: {NAVY}; border-radius: 3px; font-size: 12px; font-weight: 500;
    padding: 3px 10px; margin: 3px;
  }}

  /* Buttons */
  div[data-testid="stButton"] > button {{
    background: {NAVY}; color: {WHITE}; border: none; border-radius: 4px;
    font-weight: 600; font-size: 14px; padding: 10px 28px; width: 100%;
    transition: background 0.15s ease;
  }}
  div[data-testid="stButton"] > button:hover {{ background: {SKY}; color: {WHITE}; }}

  div[data-testid="stDownloadButton"] > button {{
    background: {WHITE}; color: {NAVY}; border: 1.5px solid {NAVY};
    border-radius: 4px; font-weight: 600; font-size: 13px;
    transition: all 0.15s ease;
  }}
  div[data-testid="stDownloadButton"] > button:hover {{ background: {NAVY}; color: {WHITE}; }}

  .result-row {{
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 14px; border-radius: 4px; margin-bottom: 6px;
    font-size: 13px; font-weight: 500;
  }}
  .result-ok  {{ background: #EAF7EE; border: 1px solid #A3D9B1; color: #1A6631; }}
  .result-err {{ background: #FEF0F0; border: 1px solid #F5B7B1; color: #922B21; }}

  /* Top bar */
  .kpmg-topbar {{
    background: {NAVY}; padding: 14px 24px; border-radius: 6px;
    display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px;
  }}
  .kpmg-topbar .title   {{ font-size: 18px; font-weight: 700; color: {WHITE}; letter-spacing: -0.01em; }}
  .kpmg-topbar .subtitle {{ font-size: 12px; color: rgba(255,255,255,0.65); margin-top: 2px; }}
  .confidential-badge {{
    background: {RED}; color: {WHITE}; font-size: 10px; font-weight: 700;
    letter-spacing: 0.12em; padding: 4px 10px; border-radius: 2px; text-transform: uppercase;
  }}

  /* AI insights */
  .ai-header {{
    background: linear-gradient(135deg, {NAVY} 0%, #0045B5 100%);
    color: {WHITE}; padding: 14px 20px; border-radius: 6px;
    margin-bottom: 16px; display: flex; align-items: center; gap: 12px;
  }}
  .ai-header .ai-title {{ font-size: 15px; font-weight: 700; }}
  .ai-header .ai-sub   {{ font-size: 11px; opacity: 0.70; margin-top: 2px; }}
  .ai-badge {{
    background: {GOLD}; color: {WHITE}; font-size: 9px; font-weight: 700;
    letter-spacing: 0.10em; padding: 3px 8px; border-radius: 2px;
    text-transform: uppercase; white-space: nowrap;
  }}

  /* Hide Streamlit chrome */
  #MainMenu, footer, header {{ visibility: hidden; }}
  .stProgress > div > div > div > div {{ background: {SKY}; }}
  span[data-baseweb="tag"] {{
    background: rgba(0,145,218,0.15) !important;
    border: 1px solid {SKY} !important;
  }}
</style>
""", unsafe_allow_html=True)


REQUIRED_COLS = _REQUIRED_COLS

# ── Session state defaults ────────────────────────────────────────────────────
for _k, _v in {
    "gen":              None,
    "chat_history":     [],
    "show_chat":        False,
    "chat_input_key":   0,
    "pe_insights":      None,
    "pe_insights_label": "",
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = os.path.join(os.path.dirname(__file__), "kpmg_logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=140)
    else:
        st.markdown("<h2 style='color:white;font-weight:900;'>KPMG</h2>", unsafe_allow_html=True)

    st.markdown("<hr class='sidebar-divider'>", unsafe_allow_html=True)
    st.markdown(
        "<p style='font-size:11px;opacity:0.6;text-transform:uppercase;"
        "letter-spacing:0.1em;margin-bottom:4px;'>Tool</p>"
        "<p style='font-size:15px;font-weight:600;margin-top:0;'>Capital Analysis<br>Statement Generator</p>",
        unsafe_allow_html=True,
    )
    st.markdown("<hr class='sidebar-divider'>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Upload investor data (.xlsx)",
        type=["xlsx"],
        help="Must contain all required columns.",
    )

    # ── PE Chat toggle ────────────────────────────────────────────────────────
    st.markdown("<hr class='sidebar-divider'>", unsafe_allow_html=True)

    chat_btn_label = "✕  Close Chat" if st.session_state["show_chat"] else "💬  PE Chat"
    if st.button(chat_btn_label, key="chat_toggle"):
        st.session_state["show_chat"] = not st.session_state["show_chat"]
        st.rerun()

    if st.session_state["show_chat"]:
        st.markdown(
            "<p style='font-size:10px;opacity:0.5;margin:4px 0 8px;'>"
            "Ask anything about PE, your portfolio, or specific investors.</p>",
            unsafe_allow_html=True,
        )

        # Render chat history
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
                "<p style='font-size:11px;opacity:0.45;font-style:italic;'>"
                "No messages yet. Ask your first question below.</p>",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        # Input + buttons
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
            if st.button("🗑", key="chat_clear_btn", use_container_width=True):
                st.session_state["chat_history"] = []
                st.rerun()

        if send_clicked and user_input.strip():
            st.session_state["chat_history"].append({"role": "user", "content": user_input.strip()})
            with st.spinner(""):
                try:
                    ctx = ""
                    if st.session_state["gen"]:
                        g = st.session_state["gen"]
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
        "<p style='font-size:10px;opacity:0.45;line-height:1.6;'>"
        "Documents generated are CONFIDENTIAL.<br>For internal use only.<br><br>"
        "© KPMG International</p>",
        unsafe_allow_html=True,
    )


# ── Main area ─────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="kpmg-topbar">
  <div>
    <div class="title">Capital Analysis Statement Generator</div>
    <div class="subtitle">Automated investor reporting — one document per investor</div>
  </div>
  <div class="confidential-badge">Confidential</div>
</div>
""", unsafe_allow_html=True)


if uploaded_file is None:
    c1, c2, c3 = st.columns(3)
    def _metric(col, label, value, sub=""):
        col.markdown(f"""
        <div class="metric-card">
          <div class="label">{label}</div>
          <div class="value">{value}</div>
          <div class="sub">{sub}</div>
        </div>""", unsafe_allow_html=True)
    _metric(c1, "Documents", "—", "Awaiting upload")
    _metric(c2, "Investors",  "—", "Awaiting upload")
    _metric(c3, "Periods",    "—", "Awaiting upload")

    st.markdown("<div class='section-header'>How it works</div>", unsafe_allow_html=True)
    st.markdown("""
Upload your investor Excel file using the sidebar. The tool will:
1. Detect all unique investors and the latest reporting period for each
2. Generate one Capital Analysis Statement per investor — **Word (.docx)** and **PDF (.pdf)**
3. Build a summary downloadable as **Excel (.xlsx)** or **CSV (.csv)**
4. Enable **AI PE Insights** and **💬 PE Chat** powered by Gemini for HNI / institutional decision-making
""")
    st.stop()


# ── Load & validate ───────────────────────────────────────────────────────────
try:
    df_raw = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Could not read file: {e}")
    st.stop()

missing = REQUIRED_COLS - set(df_raw.columns)
if missing:
    st.error(f"Missing columns: `{'`, `'.join(sorted(missing))}`")
    st.stop()

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


# ── Summary metrics ───────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
def _metric(col, label, value, sub=""):
    col.markdown(f"""
    <div class="metric-card">
      <div class="label">{label}</div>
      <div class="value">{value}</div>
      <div class="sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

_metric(c1, "Partnership", partnership[:24] + ("…" if len(partnership) > 24 else ""), "from file")
_metric(c2, "Investors",   str(n_investors),   "unique names")
_metric(c3, "Data rows",   str(len(df_raw)),   f"across {n_periods} period(s)")
_metric(c4, "As of",
        fmt_date(df_latest["TO_DATE"].max()) if n_investors else "—",
        "latest period")


# ── Investor selection ────────────────────────────────────────────────────────
st.markdown("<div class='section-header'>Investor Selection</div>", unsafe_allow_html=True)
col_sel, col_info = st.columns([2, 1])
with col_sel:
    selection_mode = st.radio(
        "Generate statements for:",
        ["All investors", "Selected investors"],
        horizontal=True, label_visibility="collapsed",
    )
with col_info:
    st.caption(f"{n_investors} investor(s) found in file.")

if selection_mode == "Selected investors":
    chosen = st.multiselect("Choose investors", options=all_investors,
                            default=all_investors, label_visibility="collapsed")
else:
    chosen = all_investors

st.markdown("".join(f"<span class='investor-pill'>{inv}</span>" for inv in chosen),
            unsafe_allow_html=True)


# ── Data preview ──────────────────────────────────────────────────────────────
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


# ── Generate ──────────────────────────────────────────────────────────────────
st.markdown("<div class='section-header'>Generate Documents</div>", unsafe_allow_html=True)

if not chosen:
    st.warning("Select at least one investor.")
    st.stop()

if st.button(f"Generate {len(chosen)} statement(s)"):
    df_selected = df_latest[df_latest["INVESTOR_NAME"].isin(chosen)].copy()

    progress_bar = st.progress(0, text="Starting…")
    results_ph   = st.empty()

    docs_in_memory: dict[str, bytes] = {}
    pdfs_in_memory: dict[str, bytes] = {}
    result_rows:    list[dict]       = []

    for idx, (_, row) in enumerate(df_selected.iterrows()):
        investor  = str(row["INVESTOR_NAME"]).strip()
        safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in investor)
        progress_bar.progress(idx / len(df_selected), text=f"Generating: {investor}")

        try:
            doc = build_document(row)
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
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
        ico = "✅" if r["ok"] else "❌"
        html_rows += (
            f"<div class='result-row {cls}'>"
            f"<span>{ico} {r['investor']}</span>"
            f"<span style='font-weight:400;opacity:0.7;font-size:12px;'>{r['msg']}</span>"
            f"</div>"
        )
    results_ph.markdown(html_rows, unsafe_allow_html=True)

    if not docs_in_memory:
        st.error("No documents were generated.")
        st.stop()

    # Build Excel + CSV
    summary_df = build_summary_excel(df_raw, df_selected)
    excel_buf  = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl", date_format="YYYY-MM-DD") as w:
        summary_df.to_excel(w, index=False, sheet_name="CapitalStatements")
    excel_buf.seek(0)
    csv_bytes = summary_df.to_csv(index=False).encode("utf-8")

    # ZIP bundles
    word_zip = io.BytesIO()
    with zipfile.ZipFile(word_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for fn, d in docs_in_memory.items(): zf.writestr(fn, d)
    word_zip.seek(0)

    pdf_zip = io.BytesIO()
    with zipfile.ZipFile(pdf_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for fn, d in pdfs_in_memory.items(): zf.writestr(fn, d)
    pdf_zip.seek(0)

    # Store everything in session state
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
    # Clear stale insights from a previous run
    st.session_state["pe_insights"]      = None
    st.session_state["pe_insights_label"] = ""


# ── Post-generation sections (persist across reruns via session_state) ─────────
gen = st.session_state.get("gen")
if gen is None:
    st.stop()

# ── Downloads ─────────────────────────────────────────────────────────────────
st.markdown("<div class='section-header'>Download</div>", unsafe_allow_html=True)

dl1, dl2, dl3, dl4 = st.columns(4)
with dl1:
    st.download_button(
        f"⬇ Word ZIP ({len(gen['docs'])})", data=gen["word_zip"],
        file_name="capital_statements_word.zip", mime="application/zip", key="dl_word_zip",
    )
with dl2:
    st.download_button(
        f"⬇ PDF ZIP ({len(gen['pdfs'])})", data=gen["pdf_zip"],
        file_name="capital_statements_pdf.zip", mime="application/zip", key="dl_pdf_zip",
    )
with dl3:
    st.download_button(
        "⬇ Summary Excel", data=gen["excel"],
        file_name="capital_statements_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_excel",
    )
with dl4:
    st.download_button(
        "⬇ Summary CSV", data=gen["csv"],
        file_name="capital_statements_summary.csv", mime="text/csv", key="dl_csv",
    )

with st.expander("Individual documents"):
    for fname, word_data in gen["docs"].items():
        label    = fname.replace("_capital_statement.docx", "").replace("_", " ")
        pdf_fname = fname.replace(".docx", ".pdf")
        ca, cb   = st.columns(2)
        with ca:
            st.download_button(f"⬇ {label} (.docx)", data=word_data, file_name=fname,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key=f"dl_word_{fname}")
        with cb:
            if pdf_fname in gen["pdfs"]:
                st.download_button(f"⬇ {label} (.pdf)", data=gen["pdfs"][pdf_fname],
                                   file_name=pdf_fname, mime="application/pdf",
                                   key=f"dl_pdf_{fname}")


# ── Output preview ────────────────────────────────────────────────────────────
st.markdown("<div class='section-header'>Output Preview</div>", unsafe_allow_html=True)
st.dataframe(gen["summary_df"], use_container_width=True, hide_index=True)


# ── AI PE Insights ────────────────────────────────────────────────────────────
st.markdown("<div class='section-header'>AI PE Insights</div>", unsafe_allow_html=True)

st.markdown(f"""
<div class="ai-header">
  <div style="flex:1;">
    <div class="ai-title">Private Equity Investment Intelligence</div>
    <div class="ai-sub">Powered by Gemini · Tailored for HNIs, Family Offices &amp; Institutional Investors</div>
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
        f"<div style='font-size:11px;font-weight:600;text-transform:uppercase;"
        f"letter-spacing:0.08em;color:{MUTED};margin-top:16px;margin-bottom:6px;'>"
        f"Report: {st.session_state['pe_insights_label']}</div>",
        unsafe_allow_html=True,
    )
    st.markdown(st.session_state["pe_insights"])
