# FinDocIQ — Technical Documentation

> **Status:** Active development
> **Owner:** yunhan088@gmail.com
> **Last updated:** 2026-06-05
> **Document version:** 0.3
> **Scope:** Production pipeline design, file map, schema, cost/accuracy findings.

---

## 1. Overview

FinDocIQ extracts financial data from Singapore bank documents (DBS, OCBC, UOB) into
structured Excel workbooks. Two separate pipelines:

- **Pillar 3 pipeline** — regulatory disclosures (~100 pages), one tab per table
- **CFO Presentation pipeline** — slide decks (20–30 slides), one tab per slide

Two AI systems involved — never conflate:
- **Gemini** — the runtime model called to read documents
- **Claude Code** — the development agent that builds the pipeline

---

## 2. Repository layout

```
FinancialParser/
├── DELIVERABLE/
│   ├── pillar3/
│   │   ├── PASS1_TOC.py              ← TOC extraction (zero API, pure Python)
│   │   ├── PASS2_Extract_to_Excel.py ← Gemini extraction → Excel (Gemini 2.5 Pro)
│   │   └── compare_excel.py          ← Cell-by-cell diff vs baseline
│   │
│   ├── CFO_Presentations/
│   │   ├── SLIDE_Extract.py          ← Full slide pipeline (Pass 0→3)
│   │   ├── chart_contracts.json      ← Per chart-type reading contracts
│   │   └── generate_story.py         ← Builds demo/story_report.html
│   │
│   ├── demo/
│   │   ├── generate_demo.py          ← Builds demo/index.html
│   │   └── index.html                ← Self-contained demo site
│   │
│   └── outputs/
│       ├── pillar3/                  ← {bank}_pillar3.xlsx, toc.json, audit/
│       └── CFO_Presentation/         ← {bank}_CFOpresentations.xlsx, audit/
│
└── MDs/
    ├── CLAUDE.md                     ← Persistent context for Claude Code sessions
    ├── DEVLOG.md                     ← Experiment log and design decisions
    └── TECHNICAL_DOCUMENTATION.md   ← This file
```

### File roles

| File | Run? | Purpose |
|---|---|---|
| `pillar3/PASS1_TOC.py` | ✅ | Production TOC extraction |
| `pillar3/PASS2_Extract_to_Excel.py` | ✅ | Production Pillar 3 extraction |
| `pillar3/compare_excel.py` | ✅ | Validate output vs baseline |
| `CFO_Presentations/SLIDE_Extract.py` | ✅ | Production slide extraction |
| `demo/generate_demo.py` | ✅ | Rebuild demo site after new extractions |

---

## 3. Pillar 3 pipeline

### 3.1 PASS1_TOC.py — deterministic TOC extraction

Zero API calls. Reads the PDF contents page, scans footers to resolve printed page refs
to physical page numbers, outputs a structured section tree.

**Output:** `outputs/pillar3/{bank}_toc.json`

**Bank-specific handling:**
- **DBS** — Part A/B/C structure, two-column TOC, cross-part anchor collision
- **OCBC** — title overflow (titles split across 2 lines), plain page numbers in footer
- **UOB** — deep subsection tree (12.1–12.11), `Page N` header format

### 3.2 PASS2_Extract_to_Excel.py — Gemini extraction

Model: **Gemini 2.5 Pro**. Input: native PDF bytes (not image). One call per page unit.

**Unit routing (deterministic Python, no API):**
- `single` — 1 section, 1 page → "Extract every table on this page"
- `multiple` — N sections share 1 page → "Read top-to-bottom, tag each table by section heading"
- `spanning` — 1 section spans multiple pages → split into ≤2-page chunks with column context carry-over

**Image fallback:** if response is empty/thin AND pdfplumber confirms table structure exists,
retry once with PNG screenshot attached alongside the PDF. Narrative pages (0 tables) do not trigger.

**Output:** `outputs/pillar3/{bank}_pillar3.xlsx` — one tab per table, 3-row header
(bank name / table ID + title / source + date).

**Key flags:**
- `--no-pause` — run all sections without stopping
- `--section 15.4` — rerun one section only
- `--force` — re-extract even if tab exists
- `--image` — force image on all calls
- `--chunk-pages N` — max pages per call (default 2)

---

## 4. CFO Presentation pipeline

### 4.1 Architecture — hybrid routing

Four passes per slide:

```
Pass 0: classify_slide()     → element_types.json  (~150 tokens, thinking_budget=0)
        ↓
        has any VISUAL_TYPES?
        ├── YES → Single-pass (text_only=True, thinking_budget=1024)
        │         image + SINGLE_PASS_PROMPT → free-form text + JSON
        └── NO  → Multi-pass (text tables only)
                  Pass 1: image + PASS1_PROMPT (thinking_budget=512) → description.txt
                  Pass 2: description only → datapoints.json (thinking_budget=0)
        ↓
Validate → arithmetic self-checks, blank labels
        ↓ (if errors: one correction pass)
Pass 3: render_to_excel() → {bank}_CFOpresentations.xlsx
```

**VISUAL_TYPES** (single-pass): `waterfall`, `stacked_bar`, `stacked_bar_with_overlay`,
`trend_line`, `kpi_grid`, `pie`, `donut_dual_ring`

**Text types** (multi-pass): `text_table`, `npa_movement_table`

### 4.2 Why single-pass for visuals

With `response_mime_type="application/json"` + `thinking_budget=0`, Gemini skips
intermediate reasoning and pattern-matches straight to JSON output. For visual tasks
(colour reading, ring period labels), this causes financial-context overrides — e.g.
assigning `sign="-"` to GP because training data says "GP = provision charge = negative",
ignoring the green bar on screen.

`text_only=True` allows Gemini to write colour observations before the JSON
("GP bar: green fill → sign=+"), making the reasoning visible and self-consistent.
`thinking_budget=1024` adds silent deliberation for ambiguous visual signals.

### 4.3 DataPoint schema

```python
class DataPoint(BaseModel):
    series:    str        # row label verbatim
    period:    str|None   # column header — ANY column, not just time periods
    value:     str        # verbatim as printed
    row_type:  str        # data|total|sub|start|end|bridge|note
    level:     int        # 0=header/total, 1=primary, 2=sub-item
    sign:      str|None   # "+" or "-" for waterfall bridge bars only
    extra_fields: dict    # always empty {} — every column is a period
    # + provenance: bank, doc_title, doc_date, slide_title, source
```

**Key design principle — `period` = column header:**
Every column on the slide becomes a `period` value. One `DataPoint` per `(series, column)` cell.
This includes time periods AND non-time columns (YoY%, QoQ%, Change, Rating, etc.).
Duplicate column headers are disambiguated with adjacent context: `"4Q25 YoY%"`, `"FY25 YoY%"`.
`extra_fields` is always empty — there are no "extra" columns, just more periods.

### 4.4 Renderer — type-aware layout

`render_element()` dispatches by `element_type`:

| Type | Layout |
|---|---|
| `waterfall` | Vertical list — one row per bridge component, S$m column, sign-prefixed values (+/-) |
| `kpi_grid` | Two-column table — Metric \| Value |
| All others | Wide pivot — unique period values become columns, series become rows |
| `other`/`none` | Skipped — commentary/bullets, no data |

**Pivot column order:** periods appear in slide reading order. Each period is immediately
followed by its own extra columns (e.g. `4Q25 | YoY% | QoQ% | FY25 | YoY%`).

### 4.5 Audit trail

Every slide saves to `outputs/CFO_Presentation/audit/{bank}_{doc}/slide_{N}/`:
- `slide_N.png` — rendered slide image
- `pass1_prompt.txt` — prompt sent
- `description.txt` — Pass 1 output (or "single-pass" placeholder)
- `response.json` — raw Gemini response (single-pass only)
- `datapoints.json` — parsed DataPoints
- `meta.json` — cost, token counts, validation errors

Resume is free — re-runs skip slides with existing `datapoints.json` unless `--force`.

### 4.6 Cost profile

| Pass | Model | Typical tokens | Cost/slide |
|---|---|---|---|
| Pass 0 classify | Gemini 2.5 Flash | ~150 out | ~$0.0001 |
| Single-pass visual | Gemini 2.5 Flash | ~1500 out + image | ~$0.004–0.012 |
| Multi-pass text | Gemini 2.5 Flash | ~3000 out total | ~$0.008–0.015 |

Full 22-slide deck: ~$0.15–0.20. DBS (30 slides): ~$0.20–0.25.

---

## 5. Design principles

- **Deterministic code > LLM for predictable structure.** Routing, rendering, cost logging — all Python. Gemini only handles content.
- **period = column, not time.** Every cell on the slide is `(series, column_header)`. No extra_fields.
- **Colour from legend, not context.** Waterfall signs come from the bar's fill colour matched to the on-slide legend. Financial label knowledge must not override visual observation.
- **text_only=True for visual passes.** Allows Gemini to write intermediate reasoning before JSON. Prevents silent pattern-matching from overriding prompt instructions.
- **Audit everything.** Every call saves prompt, response, parsed JSON. Resume is free.
- **Contracts teach method, not values.** No hardcoded bank figures, no assumed orderings.
