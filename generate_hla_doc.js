const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents, UnderlineType
} = require("docx");
const fs = require("fs");

// ── Brand tokens ──────────────────────────────────────────────────────────────
const NAVY      = "00338D";
const SKY       = "0091DA";
const LIGHT     = "EBF4FB";
const RED       = "C00000";
const WHITE     = "FFFFFF";
const DARK      = "1A1A2E";
const MUTED     = "6B7A99";
const BORDER    = "D0D9E8";
const ALT_ROW   = "F4F7FC";
const LIGHT_GRAY = "F4F6FA";

// ── Helpers ───────────────────────────────────────────────────────────────────
const FONT = "Arial";
const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function cell(text, opts = {}) {
  const {
    bold = false, color = DARK, fill = WHITE, width = 2000,
    fontSize = 20, align = AlignmentType.LEFT, italic = false,
    isHeader = false
  } = opts;

  const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER };
  const borders = { top: border, bottom: border, left: border, right: border };

  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    borders,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold: bold || isHeader, color, font: FONT, size: fontSize, italics: italic })]
    })]
  });
}

function hdrRow(cells) {
  return new TableRow({
    tableHeader: true,
    children: cells.map(([text, width]) => cell(text, { bold: true, color: WHITE, fill: NAVY, width, isHeader: true }))
  });
}

function row(cells, shade = false) {
  return new TableRow({
    children: cells.map(([text, width, opts = {}]) =>
      cell(text, { width, fill: shade ? ALT_ROW : WHITE, ...opts })
    )
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
    children: [new TextRun({ text, bold: true, color: NAVY, font: FONT, size: 32 })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 100 },
    children: [new TextRun({ text, bold: true, color: SKY, font: FONT, size: 26 })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, color: DARK, font: FONT, size: 22 })]
  });
}

function body(text, opts = {}) {
  const { bold = false, italic = false, color = DARK, spacing = { before: 60, after: 60 } } = opts;
  return new Paragraph({
    spacing,
    children: [new TextRun({ text, bold, italic, color, font: FONT, size: 22 })]
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: FONT, size: 22, color: DARK })]
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: FONT, size: 22, color: DARK })]
  });
}

function spacer(lines = 1) {
  return Array.from({ length: lines }, () =>
    new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun("")] })
  );
}

function labelValue(label, value) {
  return new Paragraph({
    spacing: { before: 50, after: 50 },
    children: [
      new TextRun({ text: `${label}: `, bold: true, font: FONT, size: 22, color: DARK }),
      new TextRun({ text: value, font: FONT, size: 22, color: DARK }),
    ]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER, space: 1 } },
    children: [new TextRun("")]
  });
}

function callout(text, fillColor = LIGHT) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      shading: { fill: fillColor, type: ShadingType.CLEAR },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 8, color: NAVY },
        bottom: noBorder, left: noBorder, right: noBorder
      },
      margins: { top: 120, bottom: 120, left: 180, right: 180 },
      children: [new Paragraph({ children: [new TextRun({ text, font: FONT, size: 22, color: DARK, italics: true })] })]
    })] })]
  });
}

// ── Cover page paragraphs ─────────────────────────────────────────────────────
function coverLine(text, opts = {}) {
  const { size = 22, bold = false, color = DARK, align = AlignmentType.LEFT, before = 80, after = 80 } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { before, after },
    children: [new TextRun({ text, bold, color, font: FONT, size })]
  });
}

// ── Main document ─────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•",
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦",
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
      ]},
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.DECIMAL, text: "%1.%2.",
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
      ]},
    ]
  },
  styles: {
    default: {
      document: { run: { font: FONT, size: 22, color: DARK } }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: FONT, color: NAVY },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: SKY },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: FONT, color: DARK },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [
    // ── SECTION 1: Cover Page ────────────────────────────────────────────────
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      children: [
        // CONFIDENTIAL badge (right-aligned)
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 40 },
          children: [new TextRun({ text: "CONFIDENTIAL", bold: true, color: RED, font: FONT, size: 18 })]
        }),

        // KPMG branding block
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [9360],
          rows: [new TableRow({ children: [new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: NAVY, type: ShadingType.CLEAR },
            borders: noBorders,
            margins: { top: 480, bottom: 480, left: 480, right: 480 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 0, after: 120 },
                children: [new TextRun({ text: "KPMG", bold: true, color: WHITE, font: FONT, size: 80 })]
              }),
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: "Digital Lighthouse | Advisory", color: "A8C4E0", font: FONT, size: 24 })]
              }),
            ]
          })] })]
        }),

        ...spacer(2),

        // Document title
        new Paragraph({
          spacing: { before: 240, after: 80 },
          children: [new TextRun({ text: "High-Level Architecture Document", bold: true, color: NAVY, font: FONT, size: 52 })]
        }),
        new Paragraph({
          spacing: { before: 0, after: 200 },
          children: [new TextRun({ text: "Capital Analysis Statement Generator", color: SKY, font: FONT, size: 32 })]
        }),

        divider(),

        // Metadata table
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2200, 7160],
          rows: [
            new TableRow({ children: [
              cell("Version",    { width: 2200, bold: true, fill: LIGHT_GRAY }),
              cell("1.0 — Draft", { width: 7160 })
            ]}),
            new TableRow({ children: [
              cell("Date",       { width: 2200, bold: true, fill: LIGHT_GRAY }),
              cell("May 2026",   { width: 7160 })
            ]}),
            new TableRow({ children: [
              cell("Author",     { width: 2200, bold: true, fill: LIGHT_GRAY }),
              cell("Samarth Madhivanan — KPMG Digital Lighthouse", { width: 7160 })
            ]}),
            new TableRow({ children: [
              cell("Classification", { width: 2200, bold: true, fill: LIGHT_GRAY }),
              cell("Confidential — Internal Use Only", { width: 7160, bold: true, color: RED })
            ]}),
            new TableRow({ children: [
              cell("Status",     { width: 2200, bold: true, fill: LIGHT_GRAY }),
              cell("Draft — Pending Review", { width: 7160 })
            ]}),
          ]
        }),

        ...spacer(2),
        new Paragraph({
          spacing: { before: 80, after: 80 },
          children: [new TextRun({ text: "This document contains proprietary and confidential information belonging to KPMG India. Distribution or reproduction without written consent is prohibited.", font: FONT, size: 18, color: MUTED, italics: true })]
        }),

        new Paragraph({ children: [new PageBreak()] }),
      ]
    },

    // ── SECTION 2: Body ──────────────────────────────────────────────────────
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY, space: 4 } },
            children: [
              new TextRun({ text: "KPMG Digital Lighthouse  |  Capital Analysis Statement Generator  |  HLA Document", font: FONT, size: 18, color: MUTED }),
            ]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY, space: 4 } },
            tabStops: [{ type: "right", position: 9360 }],
            children: [
              new TextRun({ text: "CONFIDENTIAL — Internal Use Only", font: FONT, size: 18, color: MUTED }),
              new TextRun({ text: "\tPage ", font: FONT, size: 18, color: MUTED }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: MUTED }),
              new TextRun({ text: " of ", font: FONT, size: 18, color: MUTED }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: MUTED }),
            ]
          })]
        })
      },
      children: [

        // ── TOC ─────────────────────────────────────────────────────────────
        h1("Table of Contents"),
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 1. Executive Summary ─────────────────────────────────────────────
        h1("1. Executive Summary"),
        body("The Capital Analysis Statement Generator is an automated reporting system developed by KPMG Digital Lighthouse to eliminate manual effort in producing investor-level Capital Analysis Statements. The tool processes structured investor data from Excel and generates one formatted, audit-ready Word document per investor, alongside a consolidated summary Excel workbook."),
        ...spacer(1),
        body("Prior to this solution, producing capital statements required analysts to manually extract figures from source systems, map values to Word templates, apply formatting, and distribute documents — a process prone to transcription errors, version control issues, and significant time investment per reporting cycle."),
        ...spacer(1),
        body("The system is built for extensibility. Five defined insertion points for Generative AI are embedded in the architecture, enabling capabilities such as automated narrative commentary, anomaly detection on financial figures, natural language Q&A over output documents, and intelligent column mapping for non-standard data sources."),
        ...spacer(1),
        callout("This document describes the architecture of the POC system and its planned production evolution. It is intended for technical reviewers, solution architects, and project sponsors."),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 2. Document Control ───────────────────────────────────────────────
        h1("2. Document Control"),
        h2("2.1 Version History"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1200, 1600, 2800, 3760],
          rows: [
            hdrRow([["Version", 1200], ["Date", 1600], ["Author", 2800], ["Description", 3760]]),
            row([["0.1", 1200], ["May 2026", 1600], ["Samarth Madhivanan", 2800], ["Initial draft", 3760]]),
            row([["1.0", 1200], ["May 2026", 1600], ["Samarth Madhivanan", 2800], ["Updated with GCP services, GenAI integration points, and cost estimates", 3760]], true),
          ]
        }),
        ...spacer(1),
        h2("2.2 Reviewers"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3120, 3120, 3120],
          rows: [
            hdrRow([["Name", 3120], ["Role", 3120], ["Status", 3120]]),
            row([["[Manager Name]", 3120], ["Engagement Manager", 3120], ["Pending Review", 3120]]),
            row([["[Senior Manager]", 3120], ["Solution Architect", 3120], ["Pending Review", 3120]], true),
          ]
        }),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 3. Solution Overview ──────────────────────────────────────────────
        h1("3. Solution Overview"),
        h2("3.1 Problem Statement"),
        body("Asset and Wealth Management (AWM) clients require periodic Capital Analysis Statements that summarise each investor's capital account activity, commitment status, and total estimated value. These statements are currently produced manually, involving:"),
        bullet("Manual extraction of investor data from source systems into Excel"),
        bullet("Row-by-row mapping of financial figures to a Word document template"),
        bullet("Manual formatting, currency conversion, and calculated field verification"),
        bullet("Document naming, version control, and distribution to each investor"),
        ...spacer(1),
        body("This process is repeated for every investor, every reporting period, and across multiple partnerships — resulting in high analyst hours per cycle with elevated risk of error."),
        ...spacer(1),
        h2("3.2 Solution Summary"),
        body("The Capital Analysis Statement Generator automates this end-to-end workflow. An analyst uploads a structured Excel file containing investor data. The system validates the data, selects the latest reporting period per investor, generates a formatted Word document for each investor, computes derived summary fields, and packages all outputs for download. A Streamlit-based web interface makes the tool accessible without any command-line interaction."),
        ...spacer(1),
        h2("3.3 Key Capabilities"),
        bullet("Accepts Excel input with 21 standard columns including INVESTOR_ID and CURRENCY_CODE"),
        bullet("Generates one formatted Capital Analysis Statement (.docx) per investor"),
        bullet("Produces a consolidated summary Excel workbook with derived columns: MIN_FROM_DATE, MAX_TO_DATE, RECORD_COUNT"),
        bullet("Handles multiple reporting periods per investor — always uses the latest"),
        bullet("KPMG-branded Streamlit UI with investor selection, progress tracking, and bulk download"),
        bullet("CLI interface for batch processing and pipeline integration"),
        bullet("Five defined GenAI enhancement points ready for integration"),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 4. System Architecture ────────────────────────────────────────────
        h1("4. System Architecture"),
        h2("4.1 Architecture Overview"),
        body("The system follows a layered architecture with five discrete layers: Input, Processing, GenAI, Output, and Presentation. Each layer is independently replaceable, enabling the POC to evolve into a production system without a full rebuild."),
        ...spacer(1),

        callout("Architecture diagram: refer to workflow_diagram.drawio in the project repository for the editable visual. The layers below correspond directly to the nodes in that diagram."),
        ...spacer(1),

        h2("4.2 Architecture Layers"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1800, 2200, 5360],
          rows: [
            hdrRow([["Layer", 1800], ["Component", 2200], ["Responsibility", 5360]]),
            row([["Input", 1800], ["Excel File (.xlsx)", 2200], ["Raw investor data — 21 columns, multiple periods per investor", 5360]]),
            row([["Processing", 1800], ["Python Engine", 2200], ["Data validation, period selection, computed field derivation, document assembly", 5360]], true),
            row([["GenAI", 1800], ["Gemini API / Vertex AI", 2200], ["Narrative generation, anomaly detection, column mapping, insights, Q&A", 5360]]),
            row([["Output", 1800], ["Word + Excel", 2200], ["Per-investor .docx files and consolidated summary .xlsx workbook", 5360]], true),
            row([["Presentation", 1800], ["Streamlit UI", 2200], ["Web interface for upload, investor selection, generation, and download", 5360]]),
          ]
        }),
        ...spacer(1),

        h2("4.3 Deployment Topology"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2000, 3680, 3680],
          rows: [
            hdrRow([["Aspect", 2000], ["POC (Work Computer)", 3680], ["Production (GCP)", 3680]]),
            row([["App Hosting", 2000], ["Local — streamlit run app.py", 3680], ["Cloud Run — containerised, auto-scaling", 3680]]),
            row([["AI Access", 2000], ["Gemini API key (Google AI Studio)", 3680], ["Vertex AI via service account", 3680]], true),
            row([["File Storage", 2000], ["Local filesystem", 3680], ["Cloud Storage (GCS)", 3680]]),
            row([["Data Layer", 2000], ["In-memory pandas DataFrames", 3680], ["BigQuery — multi-partnership warehouse", 3680]], true),
            row([["Auth", 2000], ["Environment variables (.env)", 3680], ["IAM + Secret Manager", 3680]]),
            row([["CI/CD", 2000], ["Manual deployment", 3680], ["Cloud Build pipeline", 3680]], true),
          ]
        }),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 5. Component Descriptions ─────────────────────────────────────────
        h1("5. Component Descriptions"),

        h2("5.1 Input Layer"),
        body("The input layer accepts a single Excel workbook (.xlsx) containing investor data. Each row represents one investor for one reporting period. Multiple rows per investor are expected; the system always selects the row with the latest TO_DATE."),
        ...spacer(1),
        h3("Required Columns (21 total)"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3200, 6160],
          rows: [
            hdrRow([["Column", 3200], ["Description", 6160]]),
            row([["FROM_DATE / TO_DATE", 3200], ["Reporting period start and end dates", 6160]]),
            row([["PARTNERSHIP_NAME", 3200], ["Name of the investment partnership", 6160]], true),
            row([["INVESTOR_NAME", 3200], ["Full name of the investor", 6160]]),
            row([["INVESTOR_ID", 3200], ["Unique investor identifier (e.g. INV-001)", 6160]], true),
            row([["CURRENCY_CODE", 3200], ["Reporting currency (USD, EUR, GBP)", 6160]]),
            row([["COMMITTED_CAPITAL", 3200], ["Total capital committed per subscription agreement", 6160]], true),
            row([["INCEPTION_TO_DATE_CONTRIBUTION / _DISTRIBUTION", 3200], ["Cumulative contributions and distributions since fund inception", 6160]]),
            row([["OPENING_YTD_NAV / CLOSING_YTD_NAV", 3200], ["Net asset value at period open and close", 6160]], true),
            row([["YTD_CONTRIBUTION / YTD_DISTRIBUTION", 3200], ["Contributions and distributions for the current period", 6160]]),
            row([["INVESTMENT_INCOME / INVESTMENT_EXPENSE", 3200], ["Gross income and expenses from investments", 6160]], true),
            row([["UNREALIZED_GAINS_LOSS / REALIZED_GAINS_LOSS", 3200], ["Mark-to-market and crystallised gains/losses", 6160]]),
            row([["MANAGEMENT_FEE / INCENTIVE_FEE", 3200], ["Manager compensation deducted from NAV", 6160]], true),
            row([["TEV / TEV_RATIO", 3200], ["Total Estimated Value and net multiple (read directly from source)", 6160]]),
          ]
        }),
        ...spacer(1),

        h2("5.2 Processing Layer"),
        body("The core Python engine handles all data validation, transformation, and document assembly. It is implemented in generate_capital_statements.py and exposes both a CLI interface and importable functions for use by the Streamlit UI."),
        ...spacer(1),
        h3("Key Processing Steps"),
        numbered("Column presence validation — aborts with a clear error if any of the 21 required columns are missing"),
        numbered("Date parsing — converts FROM_DATE and TO_DATE to pandas Timestamps for reliable sorting"),
        numbered("Period selection — groups rows by INVESTOR_NAME and selects the last (latest TO_DATE) row per investor"),
        numbered("Computed fields — derives: net investment income (INVESTMENT_INCOME minus INVESTMENT_EXPENSE), remaining capital commitment (COMMITTED_CAPITAL minus INCEPTION_TO_DATE_CONTRIBUTION)"),
        numbered("Summary aggregation — computes MIN_FROM_DATE, MAX_TO_DATE, RECORD_COUNT per investor across all input rows"),
        numbered("Document assembly — builds each Word document section using python-docx with right-aligned tab stops for financial figures"),
        numbered("Output packaging — saves .docx files and the summary Excel to the designated output directory"),
        ...spacer(1),

        h2("5.3 Output Layer"),
        h3("Per-Investor Word Document (.docx)"),
        body("Each document follows a fixed three-section structure:"),
        bullet("Summary of Capital Account — opening balance, contributions, distributions, net investment activity, fees, and closing balance"),
        bullet("Summary of Capital Commitment — committed capital, contributed to date, and remaining commitment"),
        bullet("Summary of Distributions and Valuation — total contributed, total distributed, closing balance, TEV, and net multiple"),
        ...spacer(1),
        body("All monetary values are formatted as USD with two decimal places. Negative values (fees, losses) are rendered as -$X,XXX.XX. TEV_RATIO is formatted as a decimal multiple (e.g. 1.28x). Section headers are bold and underlined. A CONFIDENTIAL label appears top-right and a standard footnote appears at the bottom of each document."),
        ...spacer(1),
        h3("Summary Excel Workbook (.xlsx)"),
        body("A single workbook named capital_statements_summary.xlsx is produced alongside the Word documents. It contains one row per investor with all 21 input columns plus three derived columns: MIN_FROM_DATE, MAX_TO_DATE, and RECORD_COUNT. Column order matches the schema shown in the source data screenshot."),
        ...spacer(1),

        h2("5.4 Presentation Layer — Streamlit UI"),
        body("The web interface is implemented in app.py using Streamlit and styled with KPMG brand tokens (navy #00338D, sky blue #0091DA). It provides:"),
        bullet("File upload via sidebar with real-time column validation"),
        bullet("Summary metrics — partnership name, investor count, row count, latest period date"),
        bullet("Investor selection — generate for all or a chosen subset"),
        bullet("Progress bar with per-investor status during generation"),
        bullet("Three download options: ZIP of all Word documents, summary Excel workbook, individual Word files"),
        bullet("Inline output Excel preview table after generation"),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 6. Technology Stack ───────────────────────────────────────────────
        h1("6. Technology Stack"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2200, 2800, 4360],
          rows: [
            hdrRow([["Category", 2200], ["Technology", 2800], ["Purpose", 4360]]),
            row([["Language", 2200], ["Python 3.10+", 2800], ["Core runtime for all processing logic", 4360]]),
            row([["Data Processing", 2200], ["pandas, openpyxl", 2800], ["Excel ingestion, DataFrame manipulation, output workbook generation", 4360]], true),
            row([["Document Generation", 2200], ["python-docx", 2800], ["Programmatic Word document assembly", 4360]]),
            row([["Web UI", 2200], ["Streamlit 1.37+", 2800], ["Browser-based interface — no frontend code required", 4360]], true),
            row([["AI — POC", 2200], ["Gemini API (google-generativeai)", 2800], ["Direct API key access to Gemini models from local machine", 4360]]),
            row([["AI — Production", 2200], ["Vertex AI (google-cloud-aiplatform)", 2800], ["Enterprise Gemini access with IAM, audit logging, and billing integration", 4360]], true),
            row([["Cloud Storage", 2200], ["Cloud Storage (GCS)", 2800], ["File input/output storage at scale", 4360]]),
            row([["Data Warehouse", 2200], ["BigQuery", 2800], ["Multi-period, multi-partnership investor data at scale", 4360]], true),
            row([["App Hosting", 2200], ["Cloud Run", 2800], ["Containerised, auto-scaling Streamlit deployment", 4360]]),
            row([["CI/CD", 2200], ["Cloud Build + Artifact Registry", 2800], ["Automated build and deployment pipeline", 4360]], true),
            row([["Secrets", 2200], ["Secret Manager", 2800], ["Secure API key and credential storage", 4360]]),
            row([["Observability", 2200], ["Cloud Logging + Monitoring", 2800], ["Error tracking, usage metrics, alerting", 4360]], true),
            row([["Containerisation", 2200], ["Docker", 2800], ["Package app for Cloud Run deployment", 4360]]),
            row([["Version Control", 2200], ["Git", 2800], ["Source code management and collaboration", 4360]], true),
            row([["IDE", 2200], ["VS Code / PyCharm", 2800], ["Development environment on work computer", 4360]]),
            row([["CLI Auth", 2200], ["gcloud CLI", 2800], ["Authenticate to GCP services from local machine", 4360]], true),
          ]
        }),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 7. GCP Services Architecture ─────────────────────────────────────
        h1("7. GCP Services Architecture"),
        h2("7.1 Services Overview"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2400, 1400, 5560],
          rows: [
            hdrRow([["Service", 2400], ["Phase", 1400], ["Role", 5560]]),
            row([["Gemini API (Google AI Studio)", 2400], ["POC", 1400], ["Direct API key access — powers all 5 GenAI features during development", 5560]]),
            row([["Vertex AI", 2400], ["Production", 1400], ["Enterprise Gemini access with full GCP integration, IAM, and audit trail", 5560]], true),
            row([["Cloud Storage (GCS)", 2400], ["Both", 1400], ["Stores input Excel files, generated .docx outputs, and document archives", 5560]]),
            row([["BigQuery", 2400], ["Production", 1400], ["Central data warehouse for investor data across all partnerships and periods", 5560]], true),
            row([["Cloud Run", 2400], ["Production", 1400], ["Hosts the Streamlit app as a containerised, auto-scaling web service", 5560]]),
            row([["Cloud Functions", 2400], ["Production", 1400], ["Event-driven triggers — e.g. auto-generate statements on new file upload to GCS", 5560]], true),
            row([["Cloud Build", 2400], ["Production", 1400], ["CI/CD pipeline — builds Docker image and deploys to Cloud Run on code push", 5560]]),
            row([["Artifact Registry", 2400], ["Production", 1400], ["Stores versioned Docker container images", 5560]], true),
            row([["Secret Manager", 2400], ["Both", 1400], ["Secure storage for Gemini API key, service account credentials, environment variables", 5560]]),
            row([["Cloud IAM", 2400], ["Both", 1400], ["Role-based access control for all GCP resources — configured before any other service", 5560]], true),
            row([["Pub/Sub", 2400], ["Production", 1400], ["Message queue for decoupled, event-driven processing at scale", 5560]]),
            row([["API Gateway", 2400], ["Production", 1400], ["Exposes the generator as a secure REST API for integration with upstream systems", 5560]], true),
            row([["Cloud Logging", 2400], ["Both", 1400], ["Centralised logs for errors, API usage, and compliance audit trail", 5560]]),
            row([["Cloud Monitoring", 2400], ["Production", 1400], ["Alerting on error rates, latency, and Vertex AI token consumption", 5560]], true),
          ]
        }),
        ...spacer(1),
        h2("7.2 IAM Roles Required"),
        body("The following IAM roles must be granted to the service account used by the application:"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3600, 5760],
          rows: [
            hdrRow([["IAM Role", 3600], ["Justification", 5760]]),
            row([["roles/aiplatform.user", 3600], ["Submit inference requests to Vertex AI (Gemini)", 5760]]),
            row([["roles/storage.objectAdmin", 3600], ["Read input files and write output files to GCS buckets", 5760]], true),
            row([["roles/bigquery.dataEditor", 3600], ["Read and write investor data in BigQuery tables", 5760]]),
            row([["roles/secretmanager.secretAccessor", 3600], ["Access API keys and credentials at runtime", 5760]], true),
            row([["roles/logging.logWriter", 3600], ["Write application logs to Cloud Logging", 5760]]),
            row([["roles/run.invoker", 3600], ["Allow authenticated users to invoke the Cloud Run service", 5760]], true),
          ]
        }),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 8. Data Flow ──────────────────────────────────────────────────────
        h1("8. Data Flow"),
        h2("8.1 End-to-End Flow"),
        numbered("Analyst uploads investors.xlsx via the Streamlit UI sidebar or passes it via CLI --input flag"),
        numbered("System validates all 21 required columns are present — aborts with a clear error message if any are missing"),
        numbered("Dates are parsed and rows are sorted by TO_DATE per investor — the latest row is selected for each unique INVESTOR_NAME"),
        numbered("For each selected investor row, the processing engine computes derived fields (net investment income, remaining commitment) and assembles the Word document section by section"),
        numbered("The GenAI layer is invoked at each of the five integration points (see Section 9) — enriching the document with narrative commentary, anomaly flags, or insights"),
        numbered("The completed Word document is serialised to bytes and stored in memory (UI) or written to disk (CLI)"),
        numbered("After all investors are processed, the summary aggregation step groups all raw input rows by INVESTOR_NAME and computes MIN_FROM_DATE, MAX_TO_DATE, and RECORD_COUNT"),
        numbered("The summary DataFrame is exported as capital_statements_summary.xlsx with all 24 columns in the defined column order"),
        numbered("The UI presents download buttons for the ZIP of all Word files, the summary Excel, and individual Word files"),
        numbered("On CLI, all files are written to the --output directory and a confirmation line is printed per investor"),
        ...spacer(1),
        h2("8.2 Data Governance"),
        bullet("No investor data is persisted server-side during the POC — all processing is in-memory"),
        bullet("In production, data written to GCS and BigQuery is encrypted at rest using Google-managed keys by default"),
        bullet("All API calls to Vertex AI are logged to Cloud Logging for audit purposes"),
        bullet("Secret Manager ensures API keys are never stored in code or environment files on shared systems"),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 9. GenAI Integration Points ───────────────────────────────────────
        h1("9. Generative AI Integration Points"),
        body("The architecture embeds five discrete GenAI integration points. Each is independently activatable — the system functions fully without any of them. They are designed as drop-in enhancements that augment the existing output without replacing the deterministic financial calculations."),
        ...spacer(1),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [300, 2100, 2400, 2400, 2160],
          rows: [
            hdrRow([["#", 300], ["Feature", 2100], ["Input", 2400], ["Output", 2400], ["Recommended Model", 2160]]),
            row([
              ["1", 300],
              ["Investor Narrative Commentary", 2100],
              ["All financial values for the investor's latest period", 2400],
              ["2-3 sentence plain-English summary appended to the Word document", 2400],
              ["Gemini 2.0 Flash", 2160]
            ]),
            row([
              ["2", 300],
              ["Anomaly Detection", 2100],
              ["Current period values + historical values for the same investor", 2400],
              ["Flagged lines with brief explanations inserted as footnotes", 2400],
              ["Gemini 2.0 Flash", 2160]
            ], true),
            row([
              ["3", 300],
              ["Intelligent Column Mapping", 2100],
              ["Column headers from a non-standard input file", 2400],
              ["Mapping from source columns to the required 21-column schema", 2400],
              ["Gemini 1.5 Pro", 2160]
            ]),
            row([
              ["4", 300],
              ["Portfolio-Level Insight Generation", 2100],
              ["Aggregated data across all investors for the period", 2400],
              ["Structured insights report highlighting top performers, concentration risk, and trend observations", 2400],
              ["Gemini 1.5 Pro", 2160]
            ], true),
            row([
              ["5", 300],
              ["Natural Language Q&A", 2100],
              ["User question + full set of generated statements as context", 2400],
              ["Direct answer grounded in the output documents — no hallucination of financial figures", 2400],
              ["Gemini 2.0 Pro", 2160]
            ]),
          ]
        }),
        ...spacer(1),
        callout("Design principle: GenAI never overwrites or replaces a financial figure. It only adds commentary, flags, or mappings alongside deterministic outputs. All computed values (NAV, TEV, fees) remain sourced exclusively from the input data."),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 10. Security & Compliance ─────────────────────────────────────────
        h1("10. Security & Compliance"),
        h2("10.1 Data Handling"),
        bullet("Input Excel files contain investor PII (names, IDs) and confidential financial data — classified as Confidential under KPMG's data classification policy"),
        bullet("POC: data remains on the analyst's local machine — no cloud transfer except API call payloads to Gemini"),
        bullet("Production: data encrypted in transit (TLS 1.3) and at rest (AES-256 via Google-managed keys)"),
        bullet("Investor data sent to Gemini API for narrative generation must be reviewed against KPMG's data residency and third-party sharing policies before production enablement"),
        ...spacer(1),
        h2("10.2 Credential Management"),
        bullet("POC: Gemini API key stored in a .env file, loaded via python-dotenv, and excluded from version control via .gitignore"),
        bullet("Production: all credentials stored in Secret Manager, accessed at runtime via service account — no credentials in code or container images"),
        bullet("Service account follows least-privilege principle — only the IAM roles listed in Section 7.2 are granted"),
        ...spacer(1),
        h2("10.3 Document Security"),
        bullet("All generated documents carry a CONFIDENTIAL header and are password-protectable via python-docx in production"),
        bullet("Distribution of generated statements must follow KPMG's document distribution policy for investor communications"),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 11. Scalability ───────────────────────────────────────────────────
        h1("11. Scalability & Performance"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2400, 3480, 3480],
          rows: [
            hdrRow([["Dimension", 2400], ["POC Baseline", 3480], ["Production Target", 3480]]),
            row([["Investors per run", 2400], ["10 – 50", 3480], ["1,000+", 3480]]),
            row([["Partnerships", 2400], ["1", 3480], ["10+", 3480]], true),
            row([["Concurrent users", 2400], ["1 (local)", 3480], ["20+ (Cloud Run auto-scaling)", 3480]]),
            row([["Document generation time", 2400], ["< 2 seconds per investor", 3480], ["< 5 seconds per investor (with GenAI)", 3480]], true),
            row([["Input file size", 2400], ["< 5 MB", 3480], ["Up to 500 MB (BigQuery ingest)", 3480]]),
            row([["Storage", 2400], ["Local disk", 3480], ["GCS — unlimited with lifecycle policies", 3480]], true),
            row([["Gemini token throughput", 2400], ["~5M tokens / month", 3480], ["100M+ tokens / month (Vertex AI quota)", 3480]]),
          ]
        }),
        ...spacer(1),
        h2("11.1 Scaling Strategy"),
        bullet("Cloud Run scales to zero when idle and scales horizontally under load — no infrastructure management required"),
        bullet("BigQuery handles petabyte-scale queries — replacing in-memory pandas processing for very large investor datasets"),
        bullet("Cloud Functions decouple generation from UI — enabling asynchronous batch processing for large runs"),
        bullet("Pub/Sub enables parallel processing of individual investors — each investor's document generation becomes an independent task"),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 12. Cost Summary ──────────────────────────────────────────────────
        h1("12. Cost Summary"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3000, 2000, 2000, 2360],
          rows: [
            hdrRow([["Service", 3000], ["POC Monthly", 2000], ["Prod Monthly", 2000], ["Prod Annual", 2360]]),
            row([["Gemini API (Google AI Studio)", 3000], ["~$10", 2000], ["$200 – $400", 2000], ["$2,400 – $4,800", 2360]]),
            row([["Vertex AI (Gemini Pro + Flash)", 3000], ["~$2", 2000], ["$1,500 – $3,000", 2000], ["$18,000 – $36,000", 2360]], true),
            row([["Cloud Run", 3000], ["—", 2000], ["$200 – $400", 2000], ["$2,400 – $4,800", 2360]]),
            row([["Cloud Storage", 3000], ["~$0.25", 2000], ["$100 – $200", 2000], ["$1,200 – $2,400", 2360]], true),
            row([["BigQuery", 3000], ["~$0.75", 2000], ["$300 – $600", 2000], ["$3,600 – $7,200", 2360]]),
            row([["Cloud Functions + Pub/Sub", 3000], ["—", 2000], ["$50 – $100", 2000], ["$600 – $1,200", 2360]], true),
            row([["Cloud Build + Artifact Registry", 3000], ["—", 2000], ["$15 – $25", 2000], ["$180 – $300", 2360]]),
            row([["Secret Manager + Logging", 3000], ["~$0.60", 2000], ["$50 – $100", 2000], ["$600 – $1,200", 2360]], true),
            row([["API Gateway", 3000], ["—", 2000], ["$15 – $25", 2000], ["$180 – $300", 2360]]),
            new TableRow({ children: [
              cell("TOTAL (with 40% buffer)", { width: 3000, bold: true, fill: NAVY, color: WHITE }),
              cell("~$60 – $100", { width: 2000, bold: true, fill: NAVY, color: WHITE }),
              cell("$2,430 – $4,850", { width: 2000, bold: true, fill: NAVY, color: WHITE }),
              cell("$29,000 – $58,000", { width: 2360, bold: true, fill: NAVY, color: WHITE }),
            ]}),
          ]
        }),
        ...spacer(1),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3120, 3120, 3120],
          rows: [
            hdrRow([["Phase", 3120], ["Budget", 3120], ["Includes", 3120]]),
            row([["POC (3 months)", 3120], ["$500", 3120], ["GCP usage, iteration overruns, developer buffer", 3120]]),
            row([["Production Year 1", 3120], ["$80,000", 3120], ["Infrastructure + additional development steps + 40% buffer", 3120]], true),
            row([["Production Year 2+", 3120], ["$55,000 / year", 3120], ["Steady-state infrastructure + maintenance + model upgrades", 3120]]),
          ]
        }),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 13. Assumptions & Constraints ─────────────────────────────────────
        h1("13. Assumptions & Constraints"),
        h2("13.1 Assumptions"),
        bullet("Input data is always provided in the defined 21-column Excel schema — non-standard columns require the GenAI column mapping feature (Integration Point 3)"),
        bullet("One row per investor per reporting period — duplicate rows for the same investor and same TO_DATE will result in one being silently dropped"),
        bullet("All monetary values in the input are in the currency specified by CURRENCY_CODE — no automatic conversion is performed"),
        bullet("TEV and TEV_RATIO are computed upstream and provided as input — the system reads them directly without recalculating"),
        bullet("The Gemini API key and GCP service account have been provisioned and are available on the work computer"),
        bullet("KPMG's data residency and third-party data sharing policies permit investor financial data to be sent to Google AI services"),
        ...spacer(1),
        h2("13.2 Constraints"),
        bullet("python-docx does not support real-time collaboration or tracked changes — documents are generated as final outputs"),
        bullet("Streamlit's file uploader has a default 200 MB limit — large input files require the CLI or BigQuery ingest path"),
        bullet("Cloud Run has a 60-minute request timeout — very large batches (1,000+ investors with GenAI) may require async processing via Cloud Functions"),
        bullet("Gemini API free tier limits apply during POC — rate limiting may occur under concurrent heavy usage"),
        ...spacer(1),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 14. Next Steps ────────────────────────────────────────────────────
        h1("14. Next Steps & Roadmap"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [600, 2400, 4560, 1800],
          rows: [
            hdrRow([["#", 600], ["Milestone", 2400], ["Description", 4560], ["Timeline", 1800]]),
            row([["1", 600], ["POC Sign-off", 2400], ["Review generated documents with engagement manager, validate accuracy against manual baseline", 4560], ["Week 1-2", 1800]]),
            row([["2", 600], ["GenAI Integration — Points 1 & 2", 2400], ["Activate investor narrative commentary and anomaly detection using Gemini API", 4560], ["Week 2-3", 1800]], true),
            row([["3", 600], ["GCP Project Setup", 2400], ["Provision GCP project, configure IAM roles, enable required APIs, set up Secret Manager", 4560], ["Week 3", 1800]]),
            row([["4", 600], ["Cloud Run Deployment", 2400], ["Containerise the Streamlit app with Docker, push to Artifact Registry, deploy to Cloud Run", 4560], ["Week 4", 1800]], true),
            row([["5", 600], ["BigQuery Integration", 2400], ["Migrate from in-memory pandas to BigQuery for multi-partnership data at scale", 4560], ["Week 5-6", 1800]]),
            row([["6", 600], ["GenAI Integration — Points 3, 4 & 5", 2400], ["Activate column mapping, portfolio insights, and Q&A chatbot via Vertex AI", 4560], ["Week 6-8", 1800]], true),
            row([["7", 600], ["Cloud Build CI/CD", 2400], ["Set up automated build and deployment pipeline triggered by Git push", 4560], ["Week 7", 1800]]),
            row([["8", 600], ["Production Hardening", 2400], ["Add monitoring, alerting, API Gateway, async processing via Pub/Sub and Cloud Functions", 4560], ["Week 8-10", 1800]], true),
            row([["9", 600], ["User Acceptance Testing", 2400], ["Full end-to-end testing with real investor data, sign-off from engagement team", 4560], ["Week 10-12", 1800]]),
            row([["10", 600], ["Production Go-Live", 2400], ["Onboard first production partnership, monitor for 4 weeks before wider rollout", 4560], ["Week 12+", 1800]], true),
          ]
        }),
        ...spacer(2),

        // ── End of document ───────────────────────────────────────────────────
        divider(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200 },
          children: [new TextRun({ text: "End of Document", font: FONT, size: 20, color: MUTED, italics: true })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "KPMG Digital Lighthouse  |  Capital Analysis Statement Generator  |  HLA v1.0  |  May 2026", font: FONT, size: 18, color: MUTED })]
        }),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("HLA_Capital_Statement_Generator.docx", buffer);
  console.log("✅ HLA_Capital_Statement_Generator.docx written successfully.");
});
