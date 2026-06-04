# CLAUDE.md

Persistent context for Claude Code sessions on FinDocIQ. Read this first every
session. **Update the "Current state", "Next tasks", and "Session log" sections
at the end of every session.**

---

## What this project is

A two-pass pipeline that extracts financial regulatory tables (Pillar 3, LCR,
RWA, NSFR) from DBS, OCBC, and UOB PDFs into a structured Excel workbook — one
tab per table. Built as a proof-of-concept for agentic document intelligence.

**Gemini is the runtime model the pipeline calls. Claude Code is the development
agent that builds the pipeline.** Never conflate the two.

---

## Active codebase (DELIVERABLE/)

The production pipeline lives in `DELIVERABLE/`. Everything else is either a
root master, an archived experiment, or documentation.

```
DELIVERABLE/
  PASS1_TOC.py              ← Run this: TOC extraction (zero API)
  PASS2_Extract_to_Excel.py ← Run this: Gemini extraction → Excel
  compare_excel.py          ← Run this: validate output vs baseline
  outputs/                  ← All outputs land here

Merge_Experiment/           ← Isolated merge detection experiments only
  extract_to_excel.py       ← Local copy of root master (GCell schema)
  test_two_pass.py          ← E-10: two-pass image→merge map experiment
  test_targeted_prompt.py   ← E-07: single-pass targeted vertical border prompt
  test_xposition_hint.py    ← E-04 variant: x-position hint injected into prompt

build_toc.py                ← Root master (PASS1 imports from here)
extract_to_excel.py         ← Root master with GCell schema (experiments import)
```

Full file map and script purposes in `MDs/TECHNICAL_DOCUMENTATION.md`.

---

## Documentation maintenance rules (follow every session)

**These are non-negotiable. Every session must maintain the docs.**

### TECHNICAL_DOCUMENTATION.md
Update whenever any of the following happen:
- A new script is added anywhere in the project — add it to the file map table
  with its purpose and "run?" status
- A script is deleted or moved — remove or update its entry
- A script's purpose changes — update its description
- A new flag or behaviour is added to PASS1 or PASS2 — update section 3
- The output schema changes — update section 3.3
- A new experiment script is added to Merge_Experiment/ — add it to section 4
- Cost profile changes significantly — update section 3.4

### DEVLOG.md
Add a new experiment entry (E-XX) whenever:
- A new approach is tried — even if it failed, especially if it failed
- A prompt change is made and tested
- A schema change is made (add/remove fields)
- A new technique is evaluated (different chunking, different model config, etc.)
- A design decision is made that future developers should understand

Each entry must include: hypothesis, method, outcome, root cause (if failed),
decision taken, and learning. See existing entries for format.

**Do not leave experiments undocumented.** An untested idea costs nothing to log;
a forgotten failed experiment costs another developer days to rediscover.

---

## Design principles (never violate)

- **Deterministic code > LLM for predictable structure.** TOC, page grouping,
  prompt selection, value typing, Excel rendering — all fixed Python. Gemini
  only handles table content.
- **Remove fields the model should not fill.** If a schema field exists, Gemini
  will use it to compress output (schema gaming). See DEVLOG E-08.
- **One string per column, always.** `GRow.values = list[str]`. No span
  encoding in production schema.
- **Audit everything.** Every Gemini call saves prompt, PDF slice, raw response,
  parsed JSON, and token counts.
- **Resume is free.** Partial runs reload from audit files — no re-billing.
- **Per-bank calibration is a feature, not a limitation.**

---

## How to run

```bash
cd DELIVERABLE
export GEMINI_API_KEY=...

# Pass 1 — TOC (zero API, ~5s)
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  PASS1_TOC.py "../OCBC_4Q25_Pillar 3.pdf"

# Pass 2 — Extract (~10-15 min, ~$1.20-1.50)
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  PASS2_Extract_to_Excel.py "../OCBC_4Q25_Pillar 3.pdf" --no-pause

# Validate
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  compare_excel.py outputs/ocbc.xlsx outputs/ocbc_baseline.xlsx --ignore-cost
```

Full run guide in `MDs/RUN_GUIDE.md`.

---

## Current state (2026-06-03)

- **Production pipeline:** DELIVERABLE/ is clean and working. OCBC and UOB
  extractions verified against baselines. DBS not yet re-run with new pipeline.
- **Schema:** `GRow.values = list[str]` — GCell reverted after E-07/E-08 showed
  schema gaming and ~80k token overhead for <2% span usage.
- **Chunk size:** default 2 pages (reduced from 4 to cap output tokens on dense
  spanning sections).
- **Open:** Merged cell detection for OCBC deferred (E-09/E-10). Planned
  approach: deterministic pdfplumber x-position mapping.
- **Merge_Experiment:** Three test scripts capturing different merge detection
  approaches. Two-pass test (E-10) confirmed failed on OCBC p97.

## Next tasks (priority order)

1. Re-run DBS with current DELIVERABLE pipeline; validate against baseline.
2. Re-run UOB with current pipeline; validate against baseline.
3. Promote clean outputs as new baselines for all three banks.
4. Implement deterministic x-position span pre-computation for OCBC (E-09).
5. Add presentation deck support (CFO decks already in root folder).

## Session log

> Append one entry per session: date, what changed, and what's next.
> Keep newest at top.

- **2026-06-03** — Major DELIVERABLE cleanup session. Reverted GCell to
  list[str] (E-07/E-08). Added 503 retry backoff, audit-based resume, stale tab
  cleanup, footnote stitching. Ran two-pass merge experiment (E-10), confirmed
  failed. Cleaned up root scripts into Merge_Experiment/. Updated
  TECHNICAL_DOCUMENTATION.md and DEVLOG.md (E-07 through E-10). OCBC and UOB
  extractions verified. Next: DBS rerun + baseline promotion.
