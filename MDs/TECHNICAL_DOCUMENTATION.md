# FinDocIQ — Technical Documentation

> **Status:** Active development
> **Owner:** yunhan088@gmail.com
> **Last updated:** 2026-06-03
> **Document version:** 0.2
> **Scope:** Production pipeline design, file map, cost/accuracy findings.

---

## 1. Overview

FinDocIQ extracts financial regulatory tables (Pillar 3, LCR, RWA, NSFR) from
Singapore bank PDF disclosures (DBS, OCBC, UOB) into a faithful Excel workbook,
one tab per table. The pipeline is two-pass: a deterministic TOC extraction pass
(zero API calls) followed by a Gemini Vision extraction pass (one API call per
page unit).

Two distinct AI systems are involved and must not be conflated:

- **Gemini** — the *runtime model* called to read PDF tables.
- **Claude Code** — the *development agent* used to build and refactor the pipeline.

### Honest framing

This is a **deterministic pipeline with a structured Gemini extraction step**,
not a multi-agent system. Page selection, unit grouping, value typing, and Excel
rendering are all deterministic Python. Only the per-table content extraction is
a model call.

---

## 2. Repository layout

```
FinancialParser/
│
├── DELIVERABLE/                  ← Production-ready scripts (ship these)
│   ├── PASS1_TOC.py              ← Pass 1: deterministic TOC extraction (zero API)
│   ├── PASS2_Extract_to_Excel.py ← Pass 2: Gemini extraction → Excel workbook
│   ├── compare_excel.py          ← Cell-by-cell diff of two Excel workbooks
│   └── outputs/                  ← All run outputs land here
│       ├── <bank>.xlsx           ← Extracted workbook (one per bank)
│       ├── <bank>_baseline.xlsx  ← Known-good reference for comparison
│       ├── <bank>_toc.json       ← TOC extracted by PASS1
│       ├── <bank>_api_usage.jsonl← Per-call token/cost log
│       ├── <bank>_cost_summary.json ← Run-level cost summary
│       ├── API_Log.xlsx          ← Shared cumulative log across all bank runs
│       └── audit/<bank>/         ← Per-call audit files (prompt, pdf, response, parsed)
│
├── Merge_Experiment/             ← Isolated experiments on merged cell detection
│   ├── extract_to_excel.py       ← Copy of root master (GCell schema intact for experiments)
│   ├── test_two_pass.py          ← Two-pass merge detection test (Pass 1: image→merge map, Pass 2: extraction)
│   ├── test_targeted_prompt.py   ← Single-pass targeted prompt test
│   └── out/                      ← Experiment outputs (audit files per run id)
│
├── MDs/                          ← Documentation
│   ├── CLAUDE.md                 ← Persistent context for Claude Code sessions
│   ├── DEVLOG.md                 ← Experiment log and design decisions
│   ├── RUN_GUIDE.md              ← Step-by-step run instructions
│   └── TECHNICAL_DOCUMENTATION.md ← This file
│
├── build_toc.py                  ← Root master: Pass 1 TOC extraction engine
├── extract_to_excel.py           ← Root master: Pass 2 extraction engine (GCell schema)
├── test_extraction.py            ← Root experiment: single-pass with x-position hint
├── test_two_pass.py              ← Root experiment: duplicate of Merge_Experiment version
│
├── DBS_4Q25_Pillar3.pdf          ← Source PDFs
├── OCBC_4Q25_Pillar 3.pdf
└── UOB_4Q25_Pillar 3.pdf
```

### File roles — what to run vs what not to touch

| File | Run? | Purpose |
|---|---|---|
| `DELIVERABLE/PASS1_TOC.py` | ✅ Run this | Production TOC extraction |
| `DELIVERABLE/PASS2_Extract_to_Excel.py` | ✅ Run this | Production extraction + Excel |
| `DELIVERABLE/compare_excel.py` | ✅ Run this | Validate output against baseline |
| `Merge_Experiment/test_two_pass.py` | ✅ For experiments | Two-pass merge detection test |
| `Merge_Experiment/test_targeted_prompt.py` | ✅ For experiments | Single-pass targeted prompt test |
| `Merge_Experiment/extract_to_excel.py` | ❌ Don't run | Import dependency for experiments only |
| `build_toc.py` (root) | ❌ Don't run | Master copy — DELIVERABLE imports from here |
| `extract_to_excel.py` (root) | ❌ Don't run | Master copy with GCell schema — used by Merge_Experiment |
| `test_extraction.py` (root) | ❌ Superseded | Earlier single-pass experiment |
| `test_two_pass.py` (root) | ❌ Duplicate | Same as Merge_Experiment version |

---

## 3. Production pipeline (DELIVERABLE/)

### 3.1 PASS1_TOC.py — deterministic TOC extraction

Zero API calls. Reads the PDF contents page, scans footers to resolve printed
page refs to physical page numbers, and outputs a structured section tree.

```bash
cd DELIVERABLE
python3 PASS1_TOC.py "../OCBC_4Q25_Pillar 3.pdf"
# → outputs/ocbc_toc.json
```

**Output:** `outputs/<bank>_toc.json` — list of sections each with `section_id`,
`title`, `start_page`, `end_page`.

**Bank-specific handling:**
- **DBS** — Part A/B/C structure, two-column TOC layout, cross-part anchor collision
- **OCBC** — title overflow (titles split across 2 lines in TOC), plain page numbers in footer
- **UOB** — deep subsection tree (12.1–12.11), `Page N` header format, `(cont'd)` heading extension

### 3.2 PASS2_Extract_to_Excel.py — Gemini extraction

Reads the TOC JSON, groups pages into units, calls Gemini once per unit, writes
one Excel tab per table.

```bash
cd DELIVERABLE
export GEMINI_API_KEY=...
python3 PASS2_Extract_to_Excel.py "../OCBC_4Q25_Pillar 3.pdf" --no-pause
# → outputs/ocbc.xlsx
```

**Key flags:**
- `--no-pause` — run all sections without stopping for review
- `--section 15.4` — rerun one section only
- `--force` — re-extract even if tab already exists
- `--dry-run` — print call plan, no API calls
- `--image` — force image attachment on all calls
- `--chunk-pages N` — max pages per Gemini call (default 2)

**Unit types and prompts:**

| Unit type | When | Prompt |
|---|---|---|
| `single` | One subsection owns one page | "extract every table on this page" |
| `multiple` | 2+ subsections share a page | "read top-to-bottom, tag each table with section_id" |
| `spanning` | One subsection spans multiple pages | "extract across all pages, combine continued tables" |
| `continuation` | Chunk 2+ of a long spanning section | "continue from previous chunk, inject column context" |

**Narrative page filtering (zero API cost):**
Before calling Gemini, each page is checked with pdfplumber:
- `page_is_narrative()` — fewer than 10 numeric tokens → skip
- `page_has_table_structure()` — fewer than 5 significant horizontal edges → skip

Both checks must pass for a page to be sent to Gemini.

**Image fallback:**
- Primary input: native PDF slice (`mime_type="application/pdf"`)
- Image retry: triggered only when `_reasonable()` fails AND `page_has_table_structure()` is True
- `_reasonable()`: fails if no tables, missing columns/rows, or no non-empty values

**Resume logic (audit-based):**
- If `outputs/audit/<bank>/<unit_id>/parsed.json` exists → reload from file, skip API call
- Group-level skip: if all tabs in a group already exist → skip entire group
- Override either with `--force`

**503 retry:** Exponential backoff — 15s, 30s, 60s — on `503 UNAVAILABLE` or `429 RESOURCE_EXHAUSTED`.

### 3.3 Output schema (GRow.values = list[str])

```python
class GColumn(BaseModel):
    group: str | None   # 2nd-level group header; null if single-level
    leaf:  str          # column header text

class GRow(BaseModel):
    row_id:   str | None  # printed line number; null for headers/footnotes
    row_type: str         # section_header | data | total | sub_header | note
    level:    int         # 0=header/total  1=primary  2=sub-item  3=rare
    parent:   str | None  # row_id of nearest ancestor; null for level 0-1
    label:    str         # verbatim row label
    values:   list[str]   # one string per column; "" for empty cells

class GTable(BaseModel):
    title:                   str
    label_header:            str
    continued_from_previous: bool
    section_id:              str        # for multiple-section pages only
    columns:                 list[GColumn]
    rows:                    list[GRow]

class Extraction(BaseModel):
    tables: list[GTable]
```

Note: `values` is `list[str]` — plain strings, one per column. `GCell` with
`span`/`merge_type` was tried and reverted (see DEVLOG E-07/E-08). Gemini must
emit exactly one string per column with no span encoding.

### 3.4 Cost logging

Every API call is logged to:
- `outputs/<bank>_api_usage.jsonl` — one record per call (timestamp, tokens, cost)
- `outputs/<bank>_cost_summary.json` — run-level totals
- `outputs/API_Log.xlsx` — shared cumulative log across all banks and runs
- `Cost` tab in the output Excel — per-call breakdown for that run

**Observed cost profile (OCBC 4Q25 Pillar 3, ~102 pages):**
- Clean single run (new code, chunk_size=2): ~$1.20–1.50
- Cost driver: output tokens from dense spanning tables (15.3, 16.1, 18.4)
- Narrative page filter saves ~3–5 calls per document
- Audit-based resume means reruns on partial failures cost $0 for completed sections

### 3.5 compare_excel.py — output validation

```bash
python3 compare_excel.py outputs/ocbc.xlsx outputs/ocbc_baseline.xlsx --ignore-cost
# exits 0 if identical, 1 if differences found
```

Options:
- `--ignore-cost` — skip the `Cost` tab (timestamps change every run)
- `--sheets-only` — only compare sheet names and order, not cell values

---

## 4. Merge_Experiment/ — experimental scripts

These scripts test different approaches to merged cell detection for OCBC tables.
They import from the **root** `extract_to_excel.py` which retains the `GCell`
schema (`span`, `merge_type`) for experimental purposes. Do not modify the root
master when running experiments.

### 4.1 test_two_pass.py

Tests a two-pass approach:
- **Pass 1** — image only (PNG at 3× scale), freeform response, focused prompt
  asking Gemini to identify every merged cell and return a merge map JSON.
- **Pass 2** — same image, structured GCell schema response, with Pass 1's merge
  map injected as a hint.

**Finding (2026-06-03, OCBC p97):** FAILED. Pass 1 returned `ncols=4` (actual is
5), missed data-spanning merges (e.g. row 11's `9,964` across 2 columns), and
only detected blank shaded regions. The wrong merge map fed into Pass 2 caused
every row to fail the span invariant. See DEVLOG E-10.

### 4.2 test_targeted_prompt.py

Tests the original targeted single-pass prompt (combined vertical border rule +
blank merge handling + full value fidelity) that was used in earlier manual
experiments. Sends image only at 3× scale with the GCell schema.

```bash
cd /Users/Qianyunhan/Desktop/FinancialParser
export GEMINI_API_KEY=...
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  Merge_Experiment/test_targeted_prompt.py \
  --pdf "OCBC_4Q25_Pillar 3.pdf" --pages 97 --id ocbc_p97_targeted
```

Output lands in `Merge_Experiment/out/<id>/`.

---

## 5. Design principles

- **Deterministic code > LLM for predictable structure.** TOC extraction, page
  grouping, prompt selection, value typing, Excel rendering — all fixed Python.
  Gemini is only asked for table content.
- **Remove fields the model should not fill.** If a schema field exists, the
  model will use it to compress output (schema gaming). `GCell` was removed from
  production schema because Gemini used `span=2` to compress 12-column tables
  into 6 cells, breaking column alignment. See DEVLOG E-08.
- **One string per column, always.** `GRow.values = list[str]` with exactly one
  entry per column. No span encoding in the production schema.
- **Audit everything.** Every Gemini call saves prompt, PDF slice, raw response,
  parsed JSON, and token usage. No call is unauditable.
- **Resume is free.** Crashed or partial runs reload from audit files — no
  re-billing for completed sections.
- **Per-bank calibration is a feature.** DBS, OCBC, UOB have legitimately
  different PDF structures. Explicit per-bank configuration in `PASS1_TOC.py`
  beats a fragile general heuristic.

---

## 6. Known open issues

- **Merged cell rendering (OCBC):** OCBC uses blank space (not borders) to
  encode column spans. Deterministic pdfplumber x-position mapping is the planned
  solution (see DEVLOG E-09). Not yet implemented.
- **503 cascades:** Gemini occasionally returns sustained 503s mid-run. The
  retry backoff (15s/30s/60s) handles transient spikes. For sustained outages,
  the audit-based resume ensures no work is lost.
- **Section 25 (Abbreviations):** Extracted as tables — not harmful but
  unnecessary. Could be skipped via a section-title filter.
- **Span validator noise:** The column-count validator (`validate_spans`) still
  fires on old audit files that contain GCell data. Post-processing or a full
  `--force` rerun clears this.

---

## 7. How to run a full bank extraction

```bash
cd /Users/Qianyunhan/Desktop/FinancialParser/DELIVERABLE

# Step 1 — TOC (zero API cost, ~5 seconds)
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  PASS1_TOC.py "../OCBC_4Q25_Pillar 3.pdf"
# → outputs/ocbc_toc.json

# Step 2 — Extract (API calls, ~10-15 min, ~$1.20-1.50)
export GEMINI_API_KEY=...
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  PASS2_Extract_to_Excel.py "../OCBC_4Q25_Pillar 3.pdf" --no-pause
# → outputs/ocbc.xlsx

# Step 3 — Validate against baseline
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  compare_excel.py outputs/ocbc.xlsx outputs/ocbc_baseline.xlsx --ignore-cost
```

**Rerunning failed sections:**
```bash
/Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
  PASS2_Extract_to_Excel.py "../OCBC_4Q25_Pillar 3.pdf" --section 16.1 --force
```

**Promoting output to new baseline after verified run:**
```bash
cp outputs/ocbc.xlsx outputs/ocbc_baseline.xlsx
```
