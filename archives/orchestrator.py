"""
orchestrator.py — end-to-end pipeline coordinator.

Replaces manually running 5 scripts in sequence. Coordinates:
  Phase 1 → Phase 2 → Extraction → Validation → Render → Structural QA

Each phase is aware of the previous phase's output. Low-confidence tables
get a critic retry. Failures are collected and reported at the end rather
than crashing the whole run.

Usage:
  python3 orchestrator.py DBS_4Q25_Pillar3.pdf
  python3 orchestrator.py DBS_4Q25_Pillar3.pdf --from extract   # skip phase1+2
  python3 orchestrator.py DBS_4Q25_Pillar3.pdf --table t016_p28 # single table
  python3 orchestrator.py DBS_4Q25_Pillar3.pdf --critic         # enable critic (off by default)
  python3 orchestrator.py DBS_4Q25_Pillar3.pdf --from render    # just re-render
"""
from __future__ import annotations
import os
import sys
import csv
import json
import time
import argparse
from dataclasses import dataclass, field
from typing import Literal

# ---------------------------------------------------------------------------
# Phase gate — ordered list, used by --from flag
# ---------------------------------------------------------------------------
PHASES = ["phase1", "phase2", "extract", "validate", "render", "qa"]


# ---------------------------------------------------------------------------
# Result tracking
# ---------------------------------------------------------------------------

@dataclass
class PhaseResult:
    phase: str
    ok: bool
    message: str = ""
    details: list[str] = field(default_factory=list)

    def print(self):
        icon = "✅" if self.ok else "❌"
        print(f"\n{icon} [{self.phase.upper()}] {self.message}")
        for d in self.details:
            print(f"   {d}")


@dataclass
class OrchestratorRun:
    pdf_path: str
    results: list[PhaseResult] = field(default_factory=list)
    tables_extracted: int = 0
    tables_failed: list[str] = field(default_factory=list)
    tables_flagged: list[str] = field(default_factory=list)

    def add(self, result: PhaseResult):
        self.results.append(result)
        result.print()

    def summary(self):
        print("\n" + "=" * 60)
        print("ORCHESTRATOR RUN SUMMARY")
        print("=" * 60)
        for r in self.results:
            icon = "✅" if r.ok else "❌"
            print(f"  {icon} {r.phase:<12} {r.message}")
        if self.tables_failed:
            print(f"\n  ⚠️  {len(self.tables_failed)} table(s) failed extraction:")
            for t in self.tables_failed:
                print(f"       {t}")
        if self.tables_flagged:
            print(f"\n  👁  {len(self.tables_flagged)} table(s) need review:")
            for t in self.tables_flagged:
                print(f"       {t}")
        all_ok = all(r.ok for r in self.results)
        print(f"\n{'🟢 All phases passed' if all_ok else '🔴 Some phases failed'}")
        print("=" * 60)


# ---------------------------------------------------------------------------
# Phase runners
# ---------------------------------------------------------------------------

def run_phase1(run: OrchestratorRun, args) -> bool:
    """TOC extraction + table map."""
    toc_path = "out/step1_toc.json"
    csv_path = "out/step2_table_map.csv"

    if os.path.exists(toc_path) and os.path.exists(csv_path) and not args.force:
        n = sum(1 for _ in open(csv_path)) - 1
        run.add(PhaseResult("phase1", True,
                            f"Skipped — outputs exist ({n} tables in manifest)",
                            ["Use --force to re-run"]))
        return True

    print("\n⏳ [PHASE1] Extracting TOC and table map...")
    try:
        from phase1 import run_step1, run_step2, main as phase1_main
        from google import genai
        client = genai.Client()

        os.makedirs("out", exist_ok=True)
        toc = run_step1(run.pdf_path, client, toc_path)
        run_step2(toc, csv_path, "out/step2_table_map.json")

        n = sum(1 for _ in open(csv_path)) - 1
        run.add(PhaseResult("phase1", True,
                            f"TOC + table map written ({n} tables)"))
        return True
    except Exception as e:
        run.add(PhaseResult("phase1", False, f"FAILED: {e}"))
        return False


def run_phase2(run: OrchestratorRun, args) -> bool:
    """Docling grid extraction."""
    docling_dir = "out/step2_docling"
    csv_path = "out/step2_table_map.csv"

    existing = len([f for f in os.listdir(docling_dir) if f.endswith(".json")]) if os.path.isdir(docling_dir) else 0
    total = sum(1 for _ in open(csv_path)) - 1 if os.path.exists(csv_path) else 0

    if existing >= total and total > 0 and not args.force:
        run.add(PhaseResult("phase2", True,
                            f"Skipped — docling grids exist ({existing}/{total})",
                            ["Use --force to re-run"]))
        return True

    print(f"\n⏳ [PHASE2] Extracting docling grids ({existing}/{total} done)...")
    try:
        from phase2 import run_phase2 as _run_phase2
        _run_phase2(
            pdf_path=run.pdf_path,
            manifest_path=csv_path,
            out_dir=docling_dir,
            target_table=args.table if args.table else None,
        )
        done = len([f for f in os.listdir(docling_dir) if f.endswith(".json")])
        run.add(PhaseResult("phase2", True, f"Docling grids written ({done} tables)"))
        return True
    except Exception as e:
        run.add(PhaseResult("phase2", False, f"FAILED: {e}"))
        return False


def run_extraction(run: OrchestratorRun, args) -> bool:
    """Per-table Gemini extraction with optional critic retry."""
    extracted_dir = "out/step3_extracted"
    csv_path = "out/step2_table_map.csv"

    existing = len([f for f in os.listdir(extracted_dir) if f.endswith(".json")]) if os.path.isdir(extracted_dir) else 0
    total = sum(1 for _ in open(csv_path)) - 1 if os.path.exists(csv_path) else 0

    if existing >= total and total > 0 and not args.force and not args.table:
        run.add(PhaseResult("extract", True,
                            f"Skipped — all extractions exist ({existing}/{total})",
                            ["Use --force to re-run or --table to re-extract one"]))
        return True

    print(f"\n⏳ [EXTRACT] Extracting tables via Gemini ({existing}/{total} done)...")
    try:
        from extract_tables import extract_tables as _extract
        _extract(
            pdf_path=run.pdf_path,
            target_id=args.table if args.table else None,
            no_critic=not args.critic,
        )

        # Collect results
        files = [f for f in os.listdir(extracted_dir) if f.endswith(".json")] if os.path.isdir(extracted_dir) else []
        failed = []
        flagged = []
        for fname in files:
            with open(os.path.join(extracted_dir, fname)) as f:
                d = json.load(f)
            chk = d.get("self_check", {})
            if str(chk.get("totals_reconcile", "")).lower() == "false":
                flagged.append(fname.replace(".json", ""))

        run.tables_extracted = len(files)
        run.tables_failed = failed
        run.tables_flagged = flagged

        details = []
        if flagged:
            details.append(f"{len(flagged)} table(s) with totals_reconcile=false: {', '.join(flagged[:5])}")
        run.add(PhaseResult("extract", True,
                            f"{len(files)}/{total} tables extracted",
                            details))
        return True
    except Exception as e:
        run.add(PhaseResult("extract", False, f"FAILED: {e}"))
        return False


def run_validation(run: OrchestratorRun, args) -> bool:
    """Arithmetic / financial-sense checks on extracted JSON."""
    extracted_dir = "out/step3_extracted"
    if not os.path.isdir(extracted_dir) or not os.listdir(extracted_dir):
        run.add(PhaseResult("validate", False, "No extracted tables found — run extraction first"))
        return False

    print("\n⏳ [VALIDATE] Running financial-sense checks...")
    try:
        from reviewer import review_all
        results = review_all(extracted_dir)
        n_ok = sum(1 for r in results if r.makes_financial_sense and not r.skipped_reason)
        n_flagged = sum(1 for r in results if r.flags)
        n_skipped = sum(1 for r in results if r.skipped_reason)

        details = []
        for r in results:
            if r.flags:
                details.append(f"  ⚠️  {r.table_id}: {len(r.flags)} flag(s)")
                for flag in r.flags[:2]:
                    details.append(f"      {flag}")

        ok = n_flagged == 0
        run.add(PhaseResult("validate", ok,
                            f"{n_ok} OK  |  {n_flagged} flagged  |  {n_skipped} skipped",
                            details[:20]))  # cap output
        return True  # validation flags don't block render
    except Exception as e:
        run.add(PhaseResult("validate", False, f"FAILED: {e}"))
        return False


def run_render(run: OrchestratorRun, args) -> bool:
    """Render Excel workbook."""
    print("\n⏳ [RENDER] Rendering Excel workbook...")
    try:
        from render_all import render_all as _render
        _render(
            extracted_dir="out/step3_extracted",
            manifest_path="out/step2_table_map.csv",
            out_path="out/step4_output.xlsx",
            pdf_path=run.pdf_path,
        )
        size_kb = os.path.getsize("out/step4_output.xlsx") // 1024
        run.add(PhaseResult("render", True,
                            f"out/step4_output.xlsx written ({size_kb} KB)"))
        return True
    except Exception as e:
        run.add(PhaseResult("render", False, f"FAILED: {e}"))
        return False


def run_qa(run: OrchestratorRun, args) -> bool:
    """Structural QA on rendered Excel."""
    xlsx = "out/step4_output.xlsx"
    if not os.path.exists(xlsx):
        run.add(PhaseResult("qa", False, "No Excel file found — run render first"))
        return False

    print("\n⏳ [QA] Structural validation of Excel...")
    try:
        from validate_excel import validate
        result = validate(xlsx)
        details = [str(i) for i in result.issues if i.severity != "INFO"]
        run.add(PhaseResult(
            "qa", result.passed,
            f"{len(result.errors)} error(s)  {len(result.warnings)} warning(s)",
            details[:20],
        ))
        return result.passed
    except Exception as e:
        run.add(PhaseResult("qa", False, f"FAILED: {e}"))
        return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

PHASE_FNS = {
    "phase1":  run_phase1,
    "phase2":  run_phase2,
    "extract": run_extraction,
    "validate": run_validation,
    "render":  run_render,
    "qa":      run_qa,
}


def main():
    p = argparse.ArgumentParser(description="FinDocIQ end-to-end orchestrator")
    p.add_argument("pdf", help="Path to source PDF")
    p.add_argument("--from", dest="from_phase", default="phase1",
                   choices=PHASES, help="Start from this phase (default: phase1)")
    p.add_argument("--to", dest="to_phase", default="qa",
                   choices=PHASES, help="Stop after this phase (default: qa)")
    p.add_argument("--table", default=None,
                   help="Run extraction/phase2 for a single table ID only")
    p.add_argument("--critic", action="store_true",
                   help="Enable critic loop (second Gemini pass per table — doubles cost, off by default)")
    p.add_argument("--force", action="store_true",
                   help="Re-run phases even if outputs already exist")
    args = p.parse_args()

    if not os.path.exists(args.pdf):
        sys.exit(f"PDF not found: {args.pdf}")

    if not os.environ.get("GEMINI_API_KEY"):
        print("!! WARNING: GEMINI_API_KEY not set — extraction will fail")

    run = OrchestratorRun(pdf_path=args.pdf)

    start_idx = PHASES.index(args.from_phase)
    stop_idx  = PHASES.index(args.to_phase)
    active_phases = PHASES[start_idx:stop_idx + 1]

    print(f"\n🚀 FinDocIQ Orchestrator")
    print(f"   PDF:    {args.pdf}")
    print(f"   Phases: {' → '.join(active_phases)}")
    if args.table:
        print(f"   Table:  {args.table}")
    if args.critic:
        print(f"   Critic: ENABLED (doubles Gemini cost)")

    t0 = time.time()
    for phase in active_phases:
        ok = PHASE_FNS[phase](run, args)
        # extraction failure blocks render; phase1 failure blocks everything
        if not ok and phase in ("phase1", "phase2"):
            print(f"\n🛑 Stopping — {phase} failed")
            break
        if not ok and phase == "extract":
            print(f"\n⚠️  Extraction had failures — continuing to render partial results")

    elapsed = time.time() - t0
    print(f"\n⏱  Total time: {elapsed:.0f}s")
    run.summary()


if __name__ == "__main__":
    main()
