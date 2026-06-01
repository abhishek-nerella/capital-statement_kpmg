#!/usr/bin/env python3
"""
Meridian Opportunities Fund  ·  HF Capital Statement Pipeline  ·  Q1 2026

Usage:
    python run_hf_pipeline.py [path/to/workbook.xlsx]

Defaults to OpenEndedFund_HedgeFund_PCAP_Q1_2026_Waterfall.xlsx in the
same directory as this script.
"""

from __future__ import annotations

import os
import sys
import traceback

from hf_pcap_engine import load_pcap
from hf_statement_generator import build_hf_docx
from hf_excel_generator import build_hf_workbook
from validation_agent import validate_hf_row
from audit_trail import start_run, log_event, close_run


def _safe_name(name: str) -> str:
    return "".join(c if c.isalnum() or c in " _-" else "_" for c in name).strip()


def main(xlsx_path: str | None = None) -> None:
    base_dir = os.path.dirname(os.path.abspath(__file__))

    if xlsx_path is None:
        xlsx_path = os.path.join(base_dir,
                                 "OpenEndedFund_HedgeFund_PCAP_Q1_2026_Waterfall.xlsx")

    if not os.path.exists(xlsx_path):
        print(f"❌  Input file not found: {xlsx_path}")
        sys.exit(1)

    out_dir = os.path.join(base_dir, "output")
    os.makedirs(out_dir, exist_ok=True)

    # ── Step 1: Load PCAP ─────────────────────────────────────────────────────
    print(f"\nLoading PCAP: {os.path.basename(xlsx_path)}")
    investors = load_pcap(xlsx_path)
    print(f"   Loaded {len(investors)} investors\n")

    # ── Summary table ─────────────────────────────────────────────────────────
    print(f"{'Investor':<38}  {'End Cap CQ':>14}  {'Gross IRR':>10}  {'AML/KYC'}")
    print("-" * 80)
    for inv in investors:
        name    = str(inv.get("INVESTOR_NAME", "—"))[:36]
        cap     = inv.get("END_CAP_CQ", 0.0) or 0.0
        irr     = inv.get("GROSS_IRR",  0.0) or 0.0
        aml     = str(inv.get("AML_KYC", "—"))
        flag    = " ⚠" if aml == "In Review" else ""
        print(f"  {name:<36}  ${cap:>13,.2f}  {irr:>9.2f}%  {aml}{flag}")
    print()

    # ── Step 2: Validate + Generate .docx per investor ────────────────────────
    run_id  = start_run(os.path.basename(xlsx_path), len(investors))
    success = 0
    fail    = 0
    skipped = 0

    print(f"Run ID: {run_id}\n")

    for inv in investors:
        raw_name = str(inv.get("INVESTOR_NAME", "UNKNOWN"))
        safe     = _safe_name(raw_name)
        out_path = os.path.join(out_dir, f"{safe}_Q1_2026.docx")

        # ── Validation ────────────────────────────────────────────────────────
        v_result = validate_hf_row(inv, run_id)
        verdict  = v_result["verdict"]

        if verdict == "INVALID":
            print(f"❌  INVALID — {raw_name}: {v_result['notes']}")
            log_event(run_id, "document_failed", raw_name,
                      {"reason": "hf_validation_invalid", "notes": v_result["notes"]})
            fail   += 1
            skipped += 1
            continue

        if verdict == "REVALUABLE":
            # Check whether the soft fail is the AML compliance flag
            aml_flagged = any(
                c.get("severity") == "compliance" and not c["pass"]
                for c in v_result["checks"]
            )
            if aml_flagged:
                print(f"   ⚠  AML/KYC In Review — {raw_name}: generating with compliance warning")
            else:
                print(f"   ⚠  Data warning — {raw_name}: {v_result['notes']}")

        # ── Generate document ─────────────────────────────────────────────────
        try:
            build_hf_docx(inv, out_path)
            log_event(run_id, "document_generated", raw_name,
                      {"file": os.path.basename(out_path), "verdict": verdict})
            print(f"✅  Generated: {os.path.basename(out_path)}")
            success += 1
        except Exception as exc:
            log_event(run_id, "document_failed", raw_name, {"reason": str(exc)})
            print(f"❌  {raw_name}: {exc}")
            if os.environ.get("DEBUG"):
                traceback.print_exc()
            fail += 1

    close_run(run_id, success, fail)
    print()

    # ── Step 3: Build companion workbook ──────────────────────────────────────
    wb_path = os.path.join(out_dir, "Meridian_HF_Companion_Q1_2026.xlsx")
    try:
        build_hf_workbook(investors, wb_path)
        print(f"✅  Workbook: {os.path.basename(wb_path)}")
    except Exception as exc:
        print(f"❌  Workbook failed: {exc}")
        if os.environ.get("DEBUG"):
            traceback.print_exc()

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n--- Done: {success} statement(s) generated, {skipped} skipped, 1 workbook ---\n")
    if fail:
        print(f"    ⚠  {fail} investor(s) failed — check output above.")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
