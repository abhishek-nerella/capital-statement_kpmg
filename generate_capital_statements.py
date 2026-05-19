"""
Capital Analysis Statement generator.
Usage: python generate_capital_statements.py --input investors.xlsx --output ./output/
"""

import argparse
import os
import sys

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Inches


# ── formatting helpers ────────────────────────────────────────────────────────

def fmt_usd(value) -> str:
    """Format a numeric value as $1,234,567.89"""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "$0.00"
    if v < 0:
        return f"-${abs(v):,.2f}"
    return f"${v:,.2f}"


def fmt_ratio(value) -> str:
    """Format TEV_RATIO as 0.72x"""
    try:
        return f"{float(value):.2f}x"
    except (TypeError, ValueError):
        return "0.00x"


def fmt_date(value) -> str:
    """Return a human-readable date string from a pandas Timestamp or string."""
    if pd.isnull(value):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%B %d, %Y")
    return str(value)


# ── docx helpers ──────────────────────────────────────────────────────────────

def _set_run_font(run, bold=False, underline=False, size_pt=11):
    run.bold = bold
    run.underline = underline
    run.font.size = Pt(size_pt)


def add_bold_underline_paragraph(doc: Document, text: str) -> None:
    """Add a paragraph whose entire text is bold + underlined."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    _set_run_font(run, bold=True, underline=True)


def add_two_col_row(doc: Document, label: str, value: str, indent: str = "") -> None:
    """Add a single-paragraph row with label left and value right (tab-separated)."""
    p = doc.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Inches(4.5), WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run(f"{indent}{label}\t{value}")
    run.font.size = Pt(11)


def add_blank(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)


def set_cell_border_bottom(cell) -> None:
    """Add a bottom border to a table cell (used for totals lines)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), "000000")
    tcBorders.append(bottom)
    tcPr.append(tcBorders)


# ── document builder ──────────────────────────────────────────────────────────

def build_document(row: pd.Series) -> Document:
    doc = Document()

    # Narrow margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # ── "CONFIDENTIAL" top-right ─────────────────────────────────────────────
    conf_para = doc.add_paragraph()
    conf_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    conf_run = conf_para.add_run("CONFIDENTIAL")
    conf_run.bold = True
    conf_run.font.size = Pt(9)
    conf_run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    # ── Header ───────────────────────────────────────────────────────────────
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title_para.add_run("Capital Analysis")
    t_run.bold = True
    t_run.font.size = Pt(16)

    sub_para = doc.add_paragraph()
    sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s_run = sub_para.add_run(f"for the period ended {fmt_date(row['TO_DATE'])}")
    s_run.font.size = Pt(12)

    # Partnership / Investor info
    add_blank(doc)
    p = doc.add_paragraph()
    p.add_run("Partnership: ").bold = True
    p.add_run(str(row["PARTNERSHIP_NAME"]))

    p2 = doc.add_paragraph()
    p2.add_run("Investor: ").bold = True
    p2.add_run(str(row["INVESTOR_NAME"]))

    p3 = doc.add_paragraph()
    p3.add_run("Investor ID: ").bold = True
    p3.add_run(str(row["INVESTOR_ID"]))

    p4 = doc.add_paragraph()
    p4.add_run("Currency: ").bold = True
    p4.add_run(str(row["CURRENCY_CODE"]))

    add_blank(doc)

    # ── Section 1: Summary of Capital Account ────────────────────────────────
    add_bold_underline_paragraph(doc, "Summary of Capital Account")
    add_blank(doc)

    add_two_col_row(
        doc,
        f"Opening Capital balance as on {fmt_date(row['FROM_DATE'])}",
        fmt_usd(row["OPENING_YTD_NAV"]),
    )
    add_two_col_row(doc, "Capital contributions during the year", fmt_usd(row["YTD_CONTRIBUTION"]))
    add_two_col_row(doc, "Distributions during the year", fmt_usd(row["YTD_DISTRIBUTION"]))

    add_blank(doc)
    p_net = doc.add_paragraph()
    p_net.add_run("Net investment activity:").font.size = Pt(11)

    net_income = float(row["INVESTMENT_INCOME"]) - float(row["INVESTMENT_EXPENSE"])
    add_two_col_row(doc, "Investment and other income", fmt_usd(net_income), indent="    ")
    add_two_col_row(
        doc,
        "Net unrealized appreciation (depreciation)",
        fmt_usd(row["UNREALIZED_GAINS_LOSS"]),
        indent="    ",
    )
    add_two_col_row(
        doc, "Net realized gain (loss)", fmt_usd(row["REALIZED_GAINS_LOSS"]), indent="    "
    )

    add_blank(doc)
    add_two_col_row(doc, "Management fees for the period", fmt_usd(row["MANAGEMENT_FEE"]))
    add_two_col_row(doc, "Incentive fees for the period", fmt_usd(row["INCENTIVE_FEE"]))

    add_blank(doc)
    add_two_col_row(
        doc,
        f"Capital balance (remaining value) at {fmt_date(row['TO_DATE'])} *",
        fmt_usd(row["CLOSING_YTD_NAV"]),
    )

    add_blank(doc)

    # ── Section 2: Summary of Capital Commitment ─────────────────────────────
    add_bold_underline_paragraph(doc, "Summary of Capital Commitment")
    add_blank(doc)

    committed = float(row["COMMITTED_CAPITAL"])
    contributed = float(row["INCEPTION_TO_DATE_CONTRIBUTION"])
    remaining = committed - contributed

    add_two_col_row(
        doc,
        "Capital commitment per subscription agreement (A)",
        fmt_usd(committed),
    )
    add_two_col_row(doc, "Capital contributed to date (B)", fmt_usd(contributed))
    add_two_col_row(doc, "Remaining capital commitment (A-B)", fmt_usd(remaining))

    add_blank(doc)

    # ── Section 3: Summary of Distributions and Valuation ───────────────────
    add_bold_underline_paragraph(doc, "Summary of Distributions and Valuation")
    add_blank(doc)

    add_two_col_row(doc, "Total capital contributed to date", fmt_usd(row["INCEPTION_TO_DATE_CONTRIBUTION"]))
    add_two_col_row(doc, "Total distributions to date", fmt_usd(row["INCEPTION_TO_DATE_DISTRIBUTION"]))
    add_two_col_row(
        doc,
        f"Capital balance (remaining value) at {fmt_date(row['TO_DATE'])} *",
        fmt_usd(row["CLOSING_YTD_NAV"]),
    )
    add_two_col_row(
        doc, "Total Estimated Value (distributions + balance)", fmt_usd(row["TEV"])
    )
    add_two_col_row(doc, "Total Estimated Value as net multiple", fmt_ratio(row["TEV_RATIO"]))

    add_blank(doc)
    add_blank(doc)

    # ── Footer footnote ───────────────────────────────────────────────────────
    footnote_para = doc.add_paragraph()
    footnote_run = footnote_para.add_run(
        "* Represents remaining value. The remaining value is based upon available "
        "information and may not represent amounts which might ultimately be realized."
    )
    footnote_run.font.size = Pt(9)
    footnote_run.italic = True

    return doc


# ── main ──────────────────────────────────────────────────────────────────────

REQUIRED_COLS = {
    "FROM_DATE", "TO_DATE", "PARTNERSHIP_NAME", "INVESTOR_NAME",
    "INVESTOR_ID", "CURRENCY_CODE",
    "COMMITTED_CAPITAL", "INCEPTION_TO_DATE_CONTRIBUTION",
    "INCEPTION_TO_DATE_DISTRIBUTION", "OPENING_YTD_NAV",
    "YTD_CONTRIBUTION", "YTD_DISTRIBUTION",
    "INVESTMENT_INCOME", "INVESTMENT_EXPENSE",
    "UNREALIZED_GAINS_LOSS", "REALIZED_GAINS_LOSS",
    "MANAGEMENT_FEE", "INCENTIVE_FEE",
    "CLOSING_YTD_NAV", "TEV", "TEV_RATIO",
}

OUTPUT_COL_ORDER = [
    "FROM_DATE", "TO_DATE", "PARTNERSHIP_NAME", "INVESTOR_NAME",
    "INVESTOR_ID", "CURRENCY_CODE",
    "COMMITTED_CAPITAL", "INCEPTION_TO_DATE_CONTRIBUTION",
    "INCEPTION_TO_DATE_DISTRIBUTION", "OPENING_YTD_NAV",
    "YTD_CONTRIBUTION", "YTD_DISTRIBUTION",
    "INVESTMENT_INCOME", "INVESTMENT_EXPENSE",
    "UNREALIZED_GAINS_LOSS", "REALIZED_GAINS_LOSS",
    "MANAGEMENT_FEE", "INCENTIVE_FEE",
    "CLOSING_YTD_NAV", "TEV", "TEV_RATIO",
    "MIN_FROM_DATE", "MAX_TO_DATE", "RECORD_COUNT",
]


def build_summary_excel(df_raw: pd.DataFrame, df_latest: pd.DataFrame) -> pd.DataFrame:
    """Return a DataFrame matching the output Excel schema (one row per investor)."""
    agg = df_raw.groupby("INVESTOR_NAME").agg(
        MIN_FROM_DATE=("FROM_DATE", "min"),
        MAX_TO_DATE=("TO_DATE", "max"),
        RECORD_COUNT=("TO_DATE", "count"),
    ).reset_index()

    out = df_latest.merge(agg, on="INVESTOR_NAME", how="left")
    return out[OUTPUT_COL_ORDER]


def main():
    parser = argparse.ArgumentParser(description="Generate Capital Analysis Statements per investor.")
    parser.add_argument("--input", required=True, help="Path to the input .xlsx file")
    parser.add_argument("--output", required=True, help="Directory to write .docx files into")
    args = parser.parse_args()

    # Load data
    try:
        df = pd.read_excel(args.input)
    except Exception as exc:
        print(f"❌ Could not read input file: {exc}", file=sys.stderr)
        sys.exit(1)

    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        print(f"❌ Input file is missing columns: {', '.join(sorted(missing))}", file=sys.stderr)
        sys.exit(1)

    # Parse dates for proper sorting
    df["TO_DATE"] = pd.to_datetime(df["TO_DATE"], errors="coerce")
    df["FROM_DATE"] = pd.to_datetime(df["FROM_DATE"], errors="coerce")

    # Keep only the latest TO_DATE row per investor
    df_latest = (
        df.sort_values("TO_DATE", ascending=True)
        .groupby("INVESTOR_NAME", as_index=False)
        .last()
    )

    # Create output directory
    os.makedirs(args.output, exist_ok=True)

    for _, row in df_latest.iterrows():
        investor = str(row["INVESTOR_NAME"]).strip()
        try:
            doc = build_document(row)
            safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in investor)
            out_path = os.path.join(args.output, f"{safe_name}_capital_statement.docx")
            doc.save(out_path)
            print(f"✅ {investor}")
        except Exception as exc:
            print(f"❌ {investor}: {exc}")

    # Write output summary Excel
    excel_path = os.path.join(args.output, "capital_statements_summary.xlsx")
    summary_df = build_summary_excel(df, df_latest)
    with pd.ExcelWriter(excel_path, engine="openpyxl", date_format="YYYY-MM-DD") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="CapitalStatements")
    print(f"\n📊 Summary Excel → {excel_path}")


if __name__ == "__main__":
    main()
