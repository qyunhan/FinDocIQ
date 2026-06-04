# FinDocIQ — Financial Document Extraction Pipeline
## Experiment Log & Technical Development Record

**Project:** FinDocIQ — Singapore Bank Regulatory Disclosure Parser  
**Scope:** DBS, OCBC, UOB — Pillar 3, LCR, Press Releases, Performance Highlights  
**Author:** Yunhan  
**Team / Context:** UOB AI Innovation Group  
**Last Updated:** May 2026  
**Status:** Active Development

### Purpose of this Document
This log records all experiments, methodology decisions, failure analyses, and structural learnings
from building a multi-stage LLM pipeline to extract tables from Singapore bank PDFs into structured
Excel workbooks. It is intended as: (1) a technical reference for future developers, (2) a context
document for AI coding agents (Claude Code), and (3) an internal record for leadership review.

---

## 1. Problem Statement

### 1.1 Background
Singapore's three major banks — DBS, OCBC, and UOB — publish regulatory disclosures (Pillar 3, LCR,
press releases, trading updates, and performance highlights) on a quarterly basis. These documents are
structured PDFs containing dense financial tables with complex layouts including multi-level row
hierarchies, merged cells, shade-coded rows, footnote references, and institution-specific formatting
conventions.

Manually extracting and normalising these tables into structured, analysis-ready Excel workbooks is
time-consuming, error-prone, and does not scale across documents and quarters.

### 1.2 Objective
Build an automated, multi-stage LLM pipeline that:
- Extracts all financial tables from any supported bank PDF into a structured Excel workbook
- Preserves formatting fidelity: bold/shading/dash conventions, hierarchy, footnotes, brand colours
- Produces consistent, reusable output schema across DBS, OCBC, and UOB
- Is codified into a reusable skill and review agent for ongoing extraction runs

### 1.3 Why This Is Hard
The core technical challenge is not simply extracting text but faithfully reconstructing table
structure and semantics in a way that is:
- **Layout-aware:** PDFs render tables differently across banks and document types. Some use native
  PDF table structures (pdfplumber-extractable); others are text-rendered with no structural grid.
- **Semantically accurate:** A greyed-out cell means 'not applicable' (None). A printed dash '-'
  means 'zero or negligible'. These must never be conflated.
- **Hierarchically correct:** Row labels have parent-child relationships (e.g. sub-items of 'of
  which:') that must be captured as metadata for downstream use.
- **Bank-specific:** Each institution has unique structural conventions — DBS uses Part A/B/C with
  page references, UOB has multi-table subsections with a specific node tree, OCBC has title-overflow
  issues in the TOC. No single generalised method works cleanly across all three without per-bank
  calibration.

### 1.4 Scope of Documents Processed

| Document Type | Bank(s) | Extraction Mode | Approx. Pages |
|---|---|---|---|
| Full Pillar 3 Disclosure | DBS (4Q25) | pdfplumber + Gemini Vision | ~92 |
| Quarterly Pillar 3 Light | DBS, OCBC, UOB (1Q26) | pdfplumber + Gemini Vision | ~15 |
| Trading Update | DBS | Gemini Vision (text-rendered) | ~6 |
| Press Release | OCBC | Gemini Vision (text-rendered) | ~8 |
| Performance Highlights | UOB | Gemini Vision (text-rendered) | ~4 |

---

## 2. Current Methodology (Production Pipeline)

The current production pipeline uses a two-pass architecture. Pass 1 constructs a structural map
of the document (TOC extraction and page range assignment). Pass 2 iterates through each
table-bearing section and runs a Gemini Vision API call with a deterministic prompt and a Pydantic
output schema.

### 2.1 Pass 1 — TOC Extraction and Page Range Assignment (`build_toc.py`)

#### 2.1.1 Overview
The goal of Pass 1 is to produce a structured section tree: a list of sections, each with a title,
section number, and assigned physical PDF page range. This drives all downstream extraction —
Pass 2 iterates over it. Zero API calls; entirely deterministic Python.

Output: `out/step1_toc.json`.

#### 2.1.2 Method: Python Script (Not LLM)
After early experiments with LLM-based TOC extraction failed (see Experiment E-01), the final
approach uses a deterministic Python script. The flow has four stages:

**1. Read the TOC pages → get section titles and their printed page refs**
The script locates the TOC (always in the first few pages) and reads it to extract entries like
`"12.2  Credit Risk Disclosure ... 27"` or `"1.1  NSFR ... A-2"`. Each entry yields a
`(section title, printed page ref)` pair. Wrapped titles (OCBC long titles split across two lines)
are rejoined before parsing.

**2. Scan the full PDF footer → resolve printed refs to physical page numbers**
The printed page ref in the TOC (`27`, `A-2`) is NOT necessarily the physical PDF page index —
cover pages, TOC pages, and bank-specific numbering schemes (DBS uses `A-1`, `A-2` etc.) all
create an offset. To resolve this, every page of the full PDF is scanned for its printed
footer/header label, building a lookup: `{"A-2": physical page 4, "27": physical page 31, ...}`.
Each section's printed ref is then mapped to its actual physical start page.

**3. Heading search fallback → for sections with no printed page ref**
Some sections (e.g. DBS subsections) have no page ref in the TOC at all. These are anchored by
scanning the body text within the parent section's page range for a line that starts with that
section number. This is the last resort before inheriting the previous section's page.

**4. Derive end_page → next section's start minus 1, then refine**
`end_page = next_section.start_page - 1`. A final heading scan then adjusts this boundary —
trimming it back if it bleeds into unrelated content, or extending it by one page if a `(cont'd)`
heading appears on the same page as the next section's start.

#### 2.1.3 Bank-Specific TOC Nuances

**DBS — Part A/B/C Structure**

DBS full Pillar 3 (~92 pages) uses a two-level section system with Part identifiers (Part A, B, C)
and page references in the format `A-2`, `B-5`, etc. The TOC extractor must:
- Detect `PART_RE = re.compile(r"^PART\s+([A-Z])\b\s*[:\-]?\s*(.+?)(?:\.{2,}|…|$)", re.I)`
  to split the document into Part groups
- Strip the hyphen and normalise: `A-2` → token `A-2` → physical page via `token_map`; the sheet
  name is derived separately as the `section_id` (e.g. `A.1.1`)
- Handle section titles that overflow to a second line in the PDF via `_merge_wrapped_lines()`,
  which appends continuation lines to the current entry until `PAGEREF_END` is satisfied
- Handle the DBS two-column TOC layout (sections 6.1/6.2/6.3 emit bare numbers first, titles later)
  via `_patch_twocol_toc()`, which matches orphan title blocks to their bare-number slots
- Avoid cross-part anchor collision: since B.1.1 and C.1.1 both have `number="1.1"`, the PDF
  outline (bookmark) lookup is skipped for Part B/C/D/E leaves via `_outline_get()`, which only
  uses bookmarks for Part A sections
- The quarterly DBS (1Q26, ~15 pages) uses simple sequential numbering — handled by the same script
  since it has no `PART` headers

**OCBC — Title Overflow Issue**

OCBC section titles occasionally exceed one line in the TOC, causing `pdfplumber` to split the title
across two text fragments. The script applies a continuation heuristic in `_merge_wrapped_lines()`:
if a line has no section number prefix (`SEC_START` does not match) and does not contain a page
reference (`PAGEREF_END` is absent), it is appended to the current entry's title string. The
`_footer_token()` function also handles the OCBC header pattern where the page number is the last
token on the first line after a year: `r"\b20\d\d\s+(\d{1,4})\s*$"`.

**UOB — Multi-Table Subsections**

UOB Pillar 3 sections frequently contain more than one sub-table per section. The section tree
captures subsection-level granularity (e.g. Section 12 contains 12.1 through 12.11). Leaf node
detection in `_is_leaf()` ensures that only the deepest numbered nodes (those with no dotted
children) become extraction targets. The `_footer_token()` function handles UOB's `Page N` format:
`r"^Page\s+(\d{1,4})\s*$"`.

The end_page refinement pass (scanning for the last physical page where the section's number
appears as a heading) handles the UOB pattern where a section's `(cont'd)` heading appears on the
same page as the next section's start — the scan extends one page past the initial `end_page`
boundary to catch this.

### 2.2 Pass 2 — Table Extraction via Gemini Vision (`extract_to_excel.py`)

#### 2.2.1 Page-to-Image Conversion
The pipeline's primary input to Gemini is a **native PDF slice** — a cut of only the relevant pages
using `cut_pdf()` (backed by `pypdfium2`). Gemini reads the PDF natively and extracts both text
and visual layout without any intermediate conversion in the happy path.

A PNG image is attached as a fallback only: `render_images()` renders pages using
`pypdfium2` at `IMAGE_SCALE = 2.0` (PDF units → pixels; this corresponds to approximately
144 DPI equivalent output since pypdfium2's default 1.0 scale is 72 DPI). Images are passed as
`image/png` alongside the PDF part.

#### 2.2.2 Deterministic Prompt Selection
Python classifies each section into a unit type based on the TOC page ranges — no LLM involved.
The unit type determines which prompt `build_prompt()` fires:

**`single`** — one subsection owns one page.
> *"You are given a SINGLE PDF page — the '[section]' subsection. It contains ONE OR MORE data
> tables. Extract EVERY distinct table as a SEPARATE entry. If the page shows two or more separate
> grids (different column headers, or a blank gap separates them) — return them as SEPARATE tables."*

**`multiple`** — two or more subsections share the same page.
> *"You are given a SINGLE PDF page containing tables belonging to MULTIPLE subsections. Read
> TOP-TO-BOTTOM. Each time you encounter a section heading, all following tables belong to THAT
> section until the next heading appears. The sections on this page are (in order): [list].
> For each table, set `section_id` to the section number it belongs to."*

**`spanning`** — one subsection spans multiple pages (up to 4 pages sent in one call).
> *"You are given PDF pages [N–M] — the '[section]' subsection. Extract EVERY distinct table
> across ALL pages. A single large table that continues across a page break should be combined
> into ONE table with `continued_from_previous=true`. Genuinely different tables (different titles
> or columns) must be SEPARATE entries."*

**Continuation prompt** — for spanning sections exceeding 4 pages, subsequent chunks receive:
> *"You are given PDF pages [N–M] — a continuation of '[section]'. CONTEXT FROM PREVIOUS CHUNK:
> the following tables were already partially extracted: ['Table Title': columns [A | B | C ...]].
> If this chunk continues any of them (same columns, no new heading), set
> `continued_from_previous=true` and emit only the new rows."*

All four prompts append the same shared rules block (`_HIER`) covering: row hierarchy levels,
value fidelity (copy verbatim, keep dashes as `"-"`), category label requirements, and merged
cell alignment (use the image to detect missing vertical borders).

Sections longer than 4 pages are chunked by `extract_unit_chunked()`. Default chunk size is 4,
overridable via `--chunk-pages N`.

#### 2.2.3 Output Schema — GTable / Extraction (Pydantic)

Every Gemini call returns output conforming to the `Extraction` Pydantic schema. The exact class
definitions, as they appear in `extract_to_excel.py`:

```python
class GColumn(BaseModel):
    group: str | None = Field(default=None,
        description="2nd-level group header spanning sub-columns; null if single-level")
    leaf:  str = Field(description="the column header text")

class GRow(BaseModel):
    row_id:   str | None = Field(default=None,
        description="printed line number EXACTLY as shown ('1','4a','14a'); "
                    "null for rows with no printed number (section headers, sub-headers, footnotes)")
    row_type: str = Field(default="data",
        description="section_header | data | total | sub_header | note")
    level:    int = Field(
        description="0=section header or grand total; 1=primary line item; "
                    "2=sub-item (indented / 'of which' / named breakdown); 3=rare")
    parent:   str | None = Field(default=None,
        description="null for level-0 and level-1 rows; for level-2+ the row_id of "
                    "the nearest row one level above")
    label:    str = Field(description="row label text, verbatim, including footnote markers")
    values:   list[GCell] = Field(default_factory=list,
        description="cell values left-to-right as GCell objects; [] for header/sub_header/note rows. "
                    "span and merge_type are always 1/'none' — pre-computed spans deferred to E-07.")

class GTable(BaseModel):
    title:        str = Field(
        description="printed table title, verbatim, including the reporting date if shown")
    label_header: str = Field(default="",
        description="header of the row-label column, e.g. 'Metric'; '' if none")
    continued_from_previous: bool = Field(default=False,
        description="true if this table is the continuation of a table that started on the "
                    "previous page (rows continue under the same columns, header NOT repeated)")
    section_id:   str = Field(default="",
        description="for multiple-section pages only: the section number this table belongs "
                    "to (e.g. '12.2'); leave '' for single-section pages")
    columns:      list[GColumn]
    rows:         list[GRow]

class Extraction(BaseModel):
    tables: list[GTable]
```

Key schema properties:
- **Cell fidelity:** Values are transcribed verbatim including thousands separators, signs, and
  dash characters (`"-"` means zero/negligible; a shaded/blank cell is `""`). Never conflated.
- **Row metadata:** `row_id` preserves source numbering; `level` encodes hierarchy (0–3);
  `parent` references the nearest ancestor `row_id`
- **Numeric encoding:** Values are kept as strings by Gemini; the deterministic `coerce()` function
  in the writer converts to Python int/float where possible, leaving percentages and dashes as-is
- **Footnotes:** Superscript markers embedded in `label` text verbatim; no separate footnotes array
  in the current schema (they appear as `row_type="note"` rows)

#### 2.2.4 Page Chunking and Continuation Prompts

An important architectural finding (see Experiment E-05) is that sending all pages of a multi-table
section to Gemini in one call — rather than page-by-page — produces better table boundary detection.
Gemini can 'see' the full context of a subsection, assign headers correctly, and avoid splitting a
table across a page boundary.

However, very long sections (more than 2 pages) are chunked by `extract_unit_chunked()` at ≤2
pages per call (default `chunk_size=2`, reduced from 4 after observing 15k–32k token outputs on
4-page calls). For chunks after the first, `build_continuation_prompt()`
constructs a context block:

```
CONTEXT FROM PREVIOUS CHUNK:
The following table(s) were already partially extracted from earlier pages of the same section.
If this chunk continues any of them (same columns, no new heading), set
continued_from_previous=true and do NOT repeat the column headers — just emit the new rows.
  - "Table Title": columns [Col1 | Col2 | Col3 ...]
```

The continuation prompt also includes the full `_HIER` rules block (column alignment, value
fidelity, category label requirements) to maintain extraction quality across chunks.

**Merge logic:** When stitching chunked tables, `extract_unit_chunked()` merges a returned table
into the previous one only when ALL of: `continued_from_previous=True`, same column count, AND
`t.title.strip() == ""` (no new title). This prevents incorrectly merging distinct date-period
tables (e.g. Dec 2025 and Jun 2025 tables) that Gemini correctly returns as separate entries with
titles even when `continued_from_previous=True`.

#### 2.2.5 Fallback Loop — Image Retry

When the first Gemini call (PDF-only, no image) returns a response that fails `_reasonable()` —
defined as: no tables, or any table missing columns or rows, or no table having any non-empty
`values` list — the pipeline triggers one retry:

```python
def _reasonable(ext: Extraction) -> bool:
    if not ext.tables:
        return False
    for t in ext.tables:
        if not t.columns or not t.rows:
            return False
        if not any(r.values for r in t.rows):
            return False
    return True
```

The retry attaches PNG images rendered at `IMAGE_SCALE = 2.0` alongside the original PDF slice.
There is no DPI blowup or row-count threshold — the trigger is purely structural (missing tables
or empty value lists), not numeric. If the retry also returns an unreasonable response, it is
used as-is and the result is flagged in the audit log (`image_used=True` in `meta.json`).

The `--image` CLI flag forces image attachment on the first call for all units.

#### 2.2.6 Cost and Performance Profile

Observed after full pipeline runs against the DBS 4Q25 Pillar 3 (~92 pages, ~30 sections):
- **Model:** `gemini-3.5-flash`
- **Pricing:** $1.50/1M input tokens, $9.00/1M output tokens
- **Typical per-section cost:** $0.05–$0.15 for single/multiple units; $0.30–$0.80 for large
  spanning sections (e.g. C.1.1 NSFR, 4 pages, ~$0.08)
- **Cost driver:** Output tokens dominate for dense tables (Gemini echoes every row/value verbatim);
  image tokens add ~$0.01–0.02 per page when the fallback fires
- **Runtime:** ~8–12 minutes per full document
- **Cost log:** Every call is appended to `out/api_usage_log.jsonl` with timestamp, label, model,
  token counts (prompt/output/thinking), and estimated cost

---

## 3. Experiment Log

Entries are ordered chronologically. Each entry records the hypothesis, method, observed outcome,
root cause analysis (where determined), and decision taken. STATUS tags: **FAILED** = approach
abandoned; **PARTIAL** = partially successful, informed current design; **ADOPTED** = incorporated
into production pipeline.

---

### E-01 — LLM-Based TOC Extraction ● FAILED / Early development

| | |
|---|---|
| **Hypothesis** | An LLM can parse a bank PDF's table of contents page and return a structured section tree without manual Python scripting |
| **Method** | Feed TOC page image(s) to Gemini with a prompt requesting a JSON array of `{section_id, title, page_number}` |
| **Outcome** | FAILED — Inconsistent results across banks. LLM frequently hallucinated section numbers, merged adjacent entries, or dropped sections entirely. Title extraction was unreliable when entries spanned two lines. |
| **Root Cause** | The TOC pages of these documents have irregular typographic layouts that confuse spatial reasoning in the LLM. The LLM had no reliable way to distinguish a continuation line from a new section entry. |
| **Decision** | Replaced with deterministic Python script (`build_toc.py`) using `pypdfium2` + `pdfplumber` + regex. The Python approach is faster, cheaper (no API call), and produces consistent output. Bank-specific calibrations (DBS Part A/B/C, OCBC title overflow) were hardcoded as rules in `_merge_wrapped_lines()`, `_patch_twocol_toc()`, and `_footer_token()`. |
| **Learning** | For structured data extraction from documents with predictable format, deterministic code is more reliable than LLM reasoning. Reserve LLM for genuinely ambiguous content. |

---

### E-02 — Docling AI Layout Model for Pre-Structuring ● FAILED / Mid development

| | |
|---|---|
| **Hypothesis** | Docling's document layout AI model can pre-parse the PDF into a structured representation (bounding boxes, table cells, text blocks), reducing the work Gemini needs to do and improving accuracy. |
| **Method** | Run Docling's layout model on each PDF page to produce a structured document object. Pass Docling's table representation alongside the page image to Gemini as additional context. |
| **Outcome** | FAILED — Gemini produced worse output when given Docling's structured representation than when given the raw image alone. Table extraction accuracy decreased and schema compliance broke (`Extraction` Pydantic validation failures increased). |
| **Root Cause** | The leading hypothesis is that Docling's output format introduced conflicting signals into the Gemini prompt — Gemini attempted to reconcile its own visual interpretation with Docling's structural representation and failed to produce clean `GTable` output. Docling's cell boundary detection was inaccurate for these dense financial layouts (especially merged cells and shaded rows), and Gemini attempted to faithfully follow incorrect cell assignments rather than reading the image directly. |
| **Decision** | Abandoned Docling integration. Reverted to pure Gemini Vision on raw page images (native PDF part via `types.Part.from_bytes(mime_type="application/pdf")`). The `phase2.py` script (Docling grid extraction) remains in the repo as a historical artifact but is not called by the production pipeline. |
| **Learning** | Providing an AI model with a pre-processed 'structured' version of its input does not always improve performance. When the pre-processor's output is imperfect, it can actively degrade the downstream model by anchoring it to incorrect structure. |

---

### E-03 — Critic Agent Validation Pass ● FAILED / Mid development

| | |
|---|---|
| **Hypothesis** | A second LLM acting as a validation agent can review each extracted table for errors, improving accuracy without requiring a human ground-truth check. Based on published multi-agent validation methodology. |
| **Method** | After each Gemini extraction, a second API call (critic agent) receives the extracted JSON and the original page image, and is asked to identify discrepancies and return a corrected version. |
| **Outcome** | FAILED — Cost catastrophic. Running the critic on every table across one full document consumed approximately $40 USD in a single overnight run (multiple full-document passes). The critic produced verbose JSON diffs and re-extractions, multiplying token consumption. |
| **Root Cause** | Without a ground-truth dataset, the critic agent has no definitive reference to validate against. It therefore produces lengthy, uncertain outputs ('possible error', 'may be incorrect') rather than confident corrections. The verbosity directly multiplied API cost. Additionally, the critic was triggered on every table including trivially simple ones where no validation was needed. |
| **Decision** | Abandoned automated critic loop. Validation is handled by deterministic structural checks (`_reasonable()`) and arithmetic checks (`reviewer.py`) post-extraction — cheaper, faster, and more auditable. |
| **Learning** | Critic/validator agents require ground truth to be meaningful. Without it, they produce noise at high cost. Deterministic rule-based validation is a better first line of defence; LLM critics should be reserved for flagged edge cases only. |

---

### E-04 — Merged Cell Detection via Spanning Width ● FAILED / Mid development

| | |
|---|---|
| **Hypothesis** | Merged cells in PDF tables can be detected by measuring whether a cell's bounding box spans more than one column width, then communicating this structure to Gemini so it produces correctly merged output. |
| **Method** | Used `pdfplumber`'s bbox coordinates to identify cells whose width exceeded the nominal single-column width by a threshold. Tagged these cells as 'merged' in the prompt to Gemini. |
| **Outcome** | FAILED — The merged-cell detector incorrectly tagged empty cells (`""`) and cells containing only a dash as spanning cells, because these cells also appeared to occupy unusual widths in `pdfplumber`'s layout analysis. This caused downstream breakage in other tables where genuine single-cell entries were incorrectly structured. |
| **Root Cause** | Empty cells and dash cells do not have content to anchor their bounding box; `pdfplumber` infers their width from surrounding layout, which can produce anomalous measurements. The spanning-width heuristic could not reliably distinguish genuine column-spanning merges from layout artefacts on empty cells. |
| **Decision** | Removed explicit merged-cell pre-detection. Gemini handles merged cell recognition from the visual image directly. The current prompt in `_HIER` instructs: "look at the IMAGE for missing vertical border lines — if a value has no vertical borders separating it from adjacent columns, that is a merged cell; place the value under the column its horizontal centre falls under, and emit `""` for every other column in the span." |
| **Learning** | PDF coordinate-based heuristics are fragile for empty/null cells. LLM visual reasoning on the raw image outperforms programmatic coordinate analysis for ambiguous layout features. |

---

### E-05 — Page-By-Page vs Multi-Page Chunking Strategy ● PARTIAL / Mid-late development

| | |
|---|---|
| **Hypothesis** | Sending PDF pages one at a time to Gemini would simplify extraction by reducing the complexity of each individual call. |
| **Method** | Tested page-by-page extraction vs. sending all pages of a section together vs. chunks of ≤4 pages. |
| **Outcome** | PARTIAL — Page-by-page extraction produced worse table boundary detection. When a table spans two pages, a page-by-page call sees an 'orphaned' bottom half of a table with no header, leading to incorrect column assignment and hierarchy errors. Multi-page (full section) calls allowed Gemini to contextualise the section as a whole, correctly identify table headers, and assign rows to the right tables. |
| **Theoretical Basis** | Large language models and vision models perform better with more context when that context is coherent and task-relevant. A single section is a coherent unit — all pages belong to the same logical structure. Splitting this context forces the model to make local decisions without global information, increasing ambiguity. The ≤4 page chunk limit is a practical constraint (token budget / image size), not a quality preference. |
| **Decision** | Adopted multi-page section-level calls as default. Maximum chunk size set to 4 pages (default in `extract_unit_chunked(chunk_size=4)`; overridable via `--chunk-pages`). Continuation prompts via `build_continuation_prompt()` used for sections exceeding 4 pages to maintain column context coherence. |
| **Learning** | Chunk size should be calibrated to semantic units, not arbitrary page counts. A section is the right unit of extraction; page-by-page is the wrong unit. |

---

### E-11 — Two-Pass Visual Read + Chart Contracts ● IMPLEMENTED / 2026-06-04

| | |
|---|---|
| **Hypothesis** | Separating perceptual work (number inventory) from structural assignment (which number belongs where) in Pass 1 reduces sign errors on waterfalls and misread values in dense charts. Injecting chart-type-specific reading contracts before description commits Gemini to the correct data model before extraction begins. |
| **Background** | Failure analysis on DBS CFO deck waterfalls showed Gemini getting sign errors on bridge bars when doing visual reading and schema filling simultaneously. Pass 1 description was underspecified — Gemini would describe structure without explicitly enumerating all visible numbers, then Pass 2 would extract from incomplete context. |
| **Method** | Three additions: (1) `chart_contracts.json` — living registry of 9 approved chart-reading contracts (waterfall, stacked_bar, stacked_bar_with_overlay, trend_line, pie, donut_dual_ring, kpi_grid, text_table, npa_movement_table). (2) Pass 0 (`classify_slide`) — cheap micro-call (~150 tokens) returning a JSON list of element types on the slide; used to select which contracts to inject. (3) Pass 1 prompt rewritten as explicit three-step process: Step 1 = inventory every visible number with location; Step 2 = describe structure; Step 3 = assign inventory to structure and verify arithmetic before proceeding. Arithmetic check in Pass 1 means sign errors are caught during description, not after extraction. |
| **New types in KNOWN_TYPES** | Added `stacked_bar_with_overlay`, `donut_dual_ring`, `npa_movement_table` — all present in real bank CFO decks but previously unclassified, causing them to fall through to Pass 1 with no contract. |
| **Contract design** | Each contract specifies DATA MODEL, HOW TO READ, ARITHMETIC CONSTRAINT, and EXTRACT AS. Waterfall contract explicitly handles the floating-bar geometry and unit ambiguity (S$m vs ppt). `stacked_bar_with_overlay` is separated from `stacked_bar` because the overlay line has no summation relationship to bars and often uses a different axis scale — without explicit rules Gemini conflates them. `donut_dual_ring` is separated from `pie` because concentric rings encode different periods spatially and Gemini conflates FY24/FY25 segments without explicit inner/outer rules. |
| **Unknown type handling** | If Pass 0 returns an invented type (e.g. `bullet_bridge`), Pass 1 is prompted to derive its own contract and write it under `DERIVED CONTRACT: {type}`. `save_derived_contract()` parses this from description.txt and appends it to `chart_contracts.json` as `status=pending_review` for human review. |
| **Resume behaviour** | Pass 0 output is cached to `element_types.json` in the slide audit folder. Re-runs with `--force` re-run Pass 0; otherwise it resumes from cache, consistent with existing audit resume pattern. |
| **Cost impact** | Pass 0: ~$0.0001/slide. Pass 1 grows by ~200–400 tokens from contract injection. Both well under $0.01/slide target. Main benefit is accuracy, not cost. |
| **Decision** | Implemented directly into `DELIVERABLE/SLIDE_Extract.py`. No separate passes/ module created — the guide's modular structure (passes/pass0_classify.py etc.) was absorbed into the single-file architecture. Rationale: the codebase is a single-file pipeline; splitting into modules adds import complexity with no benefit at this scale. |
| **Learning** | (1) Arithmetic verification in Pass 1 (description phase) is more effective than relying solely on the correction pass — it catches the root cause (wrong sign interpretation) rather than just the symptom (self_check failure). (2) Chart type granularity matters: `stacked_bar_with_overlay` and `donut_dual_ring` cannot share contracts with their simpler variants without causing misextraction. (3) The three-step inventory pattern (list all numbers → assign structure → verify arithmetic) mirrors good manual extraction practice and is transferable to other visual document types. |

---

### E-07 — GCell Schema for Span/Merge Encoding ● FAILED / 2026-06-03

| | |
|---|---|
| **Hypothesis** | Encoding `span` and `merge_type` in a `GCell(value, span, merge_type)` wrapper — replacing `list[str]` in `GRow.values` — would allow Gemini to faithfully reconstruct merged cell structure in the Excel output. |
| **Method** | Added `GCell` as the cell type in `GRow.values`. Gemini instructed to detect merged cells from missing vertical border lines in the PDF/image and emit `span=N, merge_type='aggregate'` accordingly. Excel writer called `ws.merge_cells()` on any `span>1` cell. |
| **Outcome** | FAILED. **(1) Token overhead:** GCell wrapping added ~80k output tokens (~$0.72) per full OCBC run — doubling output size — with only 1.6% of cells having `span>1`. Not worth it. **(2) Signal unavailable for OCBC:** OCBC uses shading instead of grid lines. Neither pdfplumber rect detection nor Gemini Vision can determine spans from shading alone. |
| **Decision** | Stripped span/merge instructions from prompt. Kept `GCell` in schema with `span=1` forced default — hypothesis was Gemini would default to span=1 if not asked. See E-08. |
| **Learning** | Measure token impact before committing to a schema change. Doubling output for <2% coverage is not justified. |

---

### E-08 — GCell with Forced span=1 Default ● FAILED / 2026-06-03

| | |
|---|---|
| **Hypothesis** | Keeping `GCell` in the schema but removing span instructions from the prompt would cause Gemini to always default `span=1`, saving tokens while keeping the schema intact for future use. |
| **Method** | Removed all span/merge prompt language. `GCell.span` defaulted to 1, `merge_type` to `'none'`. No instructions given — Gemini expected to emit `{"value":"X","span":1,"merge_type":"none"}` for every cell. |
| **Outcome** | FAILED — Gemini continued emitting `span>1` on wide tables unprompted. Example: section 18.1 (12 columns) returned 6 cells with `span=2` each, placing values in wrong columns. The `validate_spans` checker flagged systematic mismatches across sections 18.1, 18.2, and others. |
| **Root Cause** | When a schema field exists and can hold non-default values, Gemini uses it to compress output — emitting `span=2` to cover two columns instead of two separate cells, reducing its own token cost. This is schema gaming: the model optimises for output brevity using the tools it is given, regardless of instructions. The only way to prevent it is to remove the field entirely. |
| **Decision** | Reverted `GRow.values` fully to `list[str]`. `GCell` class removed from the codebase. Gemini now has no mechanism to express spans — it must emit exactly one string per column. |
| **Learning** | If a schema field exists, the model will use it. "Default to X" instructions do not override the model's tendency to compress. Remove fields that should not be filled; don't instruct the model to leave them at defaults. |

---

### E-09 — OCBC Merge Semantics: blank=span, dash=nil ● INSIGHT / 2026-06-03

| | |
|---|---|
| **Finding** | Manual inspection of OCBC PDF pages (e.g. p97 NSFR section, row 11) revealed that OCBC encodes merge semantics through blank space, not borders: a printed `-` means nil/zero (span=1, value exists); a completely blank cell slot means the cell is physically spanned by an adjacent value. This is the opposite of DBS/UOB which use explicit vertical grid lines. |
| **Implication** | Span detection for OCBC is a deterministic geometry problem: extract each value's x-center via pdfplumber, map to column slot via boundary map, identify contiguous blank slots adjacent to a value as its span. No vision model needed. DBS/UOB span detection is a vision problem (detect absent vertical borders) — already partially supported by the existing `_HIER` prompt rules when images are attached. |
| **Why not built now** | The primary value of the pipeline is correct values, not correct merge rendering. A merged cell with wrong span doesn't change the extracted number — it only affects Excel visual layout. Deferring until presentation deck support is added, at which point multiple formats can be designed against together. |
| **Future path** | Per-bank span pre-computation: (1) OCBC — pdfplumber x-position mapping, inject per-row span hint before Gemini call. (2) DBS/UOB — grid-line detection via pdfplumber edges, already partially feasible. (3) Presentation decks — unknown format, assess when encountered. |
| **Learning** | Merge conventions are bank-specific and cannot be unified under one detection method. Design per-format solutions. Ship correct values first; merge rendering is polish. |

---

### E-10 — Two-Pass and Targeted Prompt Merge Detection ● FAILED / 2026-06-03

| | |
|---|---|
| **Hypothesis** | A targeted prompt sending the page image to Gemini — combined with a vertical border rule and mathematical validation — would correctly identify OCBC merged cells, either as a standalone single pass or as Pass 1 of a two-pass architecture. |
| **Background** | Manual experiments in the Gemini chat UI showed that a targeted prompt with the vertical border rule did return correct span information for OCBC p97 (e.g. row 11 `9,964` with `span=3`, row 14 blank region `span=4`). The hypothesis was that this could be replicated via the API. |
| **Experiment 1 — Single-pass targeted prompt (test_targeted_prompt.py)** | Sent page 97 image at 3× scale with the targeted vertical border rule prompt. First attempt used `build_config()` with `response_schema=Extraction` and `max_output_tokens=65536` — response truncated at 65k tokens (JSON parse error) because GCell wrapping doubled output size. Second attempt used freeform JSON config with no schema and `max_output_tokens=131072` — response returned only blank merges (shaded empty regions), completely missing data value merges like row 11's `9,964 span=3`. The chat result could not be replicated via API. |
| **Experiment 2 — Two-pass (test_two_pass.py)** | Pass 1: image only, freeform prompt asking for merge map JSON. Pass 2: image + merge map injected as hint, GCell schema. Pass 1 returned `ncols=4` (actual is 5), missed all data value merges, only found blank regions. The wrong column count cascaded into every row failing the span invariant in Pass 2. |
| **Root cause** | Four compounding failures: (1) **OCBC has no visible grid lines** — the vertical border rule fires on nothing in data rows. Gemini can detect blank shaded regions but not value merges where a number sits across two columns with no divider. (2) **Chat vs API model difference** — Gemini chat confirmed it uses Gemini 1.5 Pro with contextual memory from corrections, while the API was calling `gemini-2.5-flash` with no session memory. (3) **Image resolution** — the API was not setting `media_resolution` on the image part, meaning Gemini may have processed a downsampled version. The fix is `types.Part.from_bytes(..., media_resolution=types.PartMediaResolutionLevel.MEDIA_RESOLUTION_HIGH)` — now applied to all image calls in both the merge experiments and `SLIDE_Extract.py`. (4) **Schema truncation** — the `response_schema=Extraction` constraint forces GCell wrapping on every cell, doubling output tokens and hitting the cap before the full table is returned. |
| **Decision** | All image-based merge detection approaches for OCBC abandoned. The correct solution is deterministic pdfplumber x-position mapping (E-09) — free, no API call, bank-specific. Production schema reverted to `list[str]` (E-08). |
| **Experiment 3 — 1.5-pro with MEDIA_RESOLUTION_HIGH (test_targeted_prompt.py --model gemini-1.5-pro)** | Retried with `gemini-1.5-pro` (the model Gemini chat confirmed it uses) and `media_resolution=MEDIA_RESOLUTION_HIGH` set on the image part. Result: still only blank merges detected. Row 11's `9,964` spanning cols 2-3 still not identified. Same pattern as 2.5-flash — blank shaded regions detected, data value merges missed. |
| **Final conclusion** | Vision-based span detection for OCBC data rows is not reliably achievable regardless of model version or image resolution. The signal is absent: OCBC uses whitespace/shading with no visible border difference between a single-column value and a multi-column spanning value. No vision model can infer span from a value's horizontal position without knowing the exact column boundary coordinates — which is precisely what pdfplumber provides for free. |
| **Experiment scripts preserved** | `Merge_Experiment/test_targeted_prompt.py` and `Merge_Experiment/test_two_pass.py` kept for reference. Both now support `--model` flag and `MEDIA_RESOLUTION_HIGH`. |
| **Learning** | (1) A result that works in the chat UI may not replicate via the API — model versions, system prompts, and contextual corrections all differ. Always validate on the API before building a pipeline dependency. (2) For OCBC specifically: no model version or resolution setting fixes a missing visual signal. The correct solution is deterministic geometry (pdfplumber x-positions), not better vision. (3) `media_resolution=MEDIA_RESOLUTION_HIGH` is now set on all image calls as a baseline best practice regardless. |

---

### E-06 — UOB Multi-Table Subsection Structural Error ● PARTIAL / Late development

| | |
|---|---|
| **Symptom** | Tables extracted from UOB Pillar 3 were being assigned to incorrect subsections. Sections with more than one sub-table had their tables either merged incorrectly or dropped. Section 12.4's tab received tables that belonged to section 12.2. |
| **Root Cause Analysis** | The routing logic (`route_tables()`) relied on title-similarity matching between Gemini-returned table titles and leaf section titles. On multi-section pages (e.g. UOB p28 with 12.2/12.3/12.4 all visible), Gemini was assigning `section_id` fields inconsistently, and the title-match fallback was pulling the second table of section 12.2 into section 12.4 because 12.4's title scored higher than 12.2's second occurrence. |
| **Tree Structure Context** | The section tree is represented as leaf nodes at the subsection level (e.g. `12.2`, `12.3`, `12.4`). On a shared page, Gemini must read top-to-bottom and assign `section_id` to each table based on the most recently seen section heading. The extraction call fires at the group level (all leaves on the same pages together), not individually per leaf. |
| **Fix Applied** | Added `section_id` field to `GTable` schema. The `multiple`-unit prompt now instructs: "read the page TOP-TO-BOTTOM; each time you encounter a section heading, all following tables belong to THAT section — until the next section heading appears; set `section_id` to the section NUMBER (e.g. '12.2')." The `route_tables()` function uses `section_id` as primary routing, with title-score and overflow as fallbacks. Overflow tables go to the last matched leaf in reading order (not the next free leaf). |
| **Decision** | Fix adopted into production. Added `leaves_tagged` metadata to units so continuation sections are flagged in the prompt. Verified that UOB section 12 (12.1–12.11, ~20 pages) now routes all tables to correct tabs. |
| **Learning** | When a document's page contains multiple section headings, the model must be told explicitly to treat section headings as routing boundaries. Title-similarity matching alone is insufficient when table titles are short or similar across sections. |

---

## 4. Generalisation Constraints

A recurring question from stakeholders is: why can't the pipeline work identically across all
banks without per-bank configuration? This section documents the structural reasons.

### 4.1 Why a Single Method Does Not Generalise

| Dimension | DBS | OCBC | UOB |
|---|---|---|---|
| Section structure | Part A/B/C with page refs (`A-2`) | Sequential numbers (1, 2, 3) | Sequential numbers; multi-table subsections |
| TOC title overflow | Rare | Common — titles split across 2 lines | Rare |
| Extraction mode | pypdfium2 + pdfplumber + Gemini | pypdfium2 + pdfplumber + Gemini | pypdfium2 + pdfplumber + Gemini |
| Sub-table depth | Mostly flat (1 table per section) | Mostly flat | Deep (12.1, 12.2, … 12.11 per section) |
| Brand colour | Red — `CC0000` | Red — `CC0000` | Blue — `1B6EC2` |
| Page ref format | `A-2` → physical page via `token_map` | Plain page numbers in footer | `Page N` in header |
| Cross-part anchor collision | B.1.1 and C.1.1 both `number="1.1"` | N/A | N/A |

### 4.2 Accepted Approach
Rather than attempting full generalisation, the current approach:
- Accepts per-bank configuration as a design feature, not a limitation
- Encodes bank-specific rules as named constants and heuristic branches in `build_toc.py`
  (`BANKS` dict, `_footer_token()`, `_outline_get()`, `_patch_twocol_toc()`)
- Uses auto-detection of bank identity (`detect_bank()` in `extract_to_excel.py`) to set
  `INSTITUTION`, `BRAND_COLOUR`, and `DOC_DATE` at runtime
- Maintains a shared output schema (`GTable`, `Extraction`, Excel workbook structure) so that
  despite different extraction paths, all outputs are structurally identical

### 4.3 Future Generalisation Path
Options under consideration:
- Training a lightweight classifier on document type (PDF metadata + first-page text features)
  to auto-select the bank-specific TOC extraction branch
- Building a document-type registry (YAML) so new bank formats can be added without code changes
- Extending the `BANKS` dict to cover additional institutions (e.g. Standard Chartered SG,
  Maybank) with their own `match`, `brand`, and footer-token patterns

---

## 5. Key Learnings & Design Principles

Consolidated from all experiments and iterations:

- **Deterministic code > LLM for predictable structure.** TOC extraction, page range assignment,
  and branch logic (which prompt to fire) are all better handled by deterministic Python than LLM
  calls. Reserve LLM for genuine ambiguity — i.e., the actual table content.

- **Semantic unit > arbitrary chunk size.** Section-level page grouping (all pages of one section
  together) outperforms page-by-page extraction because it preserves the coherent context the LLM
  needs for correct table boundary detection.

- **Schema validity ≠ semantic correctness.** Pydantic validates structure, not meaning. A response
  can pass schema validation while containing incorrect values (e.g. a merged cell placed in the
  wrong column). Prompt-level rules (`_HIER` merged-cell alignment block) and post-extraction
  arithmetic checks (`reviewer.py`) are needed as a second layer.

- **Critic agents require ground truth.** Without a labelled reference dataset, an LLM critic
  produces uncertain, verbose, expensive outputs. Deterministic validation is a better first defence.

- **Pre-processing can degrade downstream LLM accuracy.** The Docling experiment showed that giving
  a model a pre-structured representation of its input can hurt performance if that pre-structure is
  imperfect. Raw PDF input (native `application/pdf` part) was superior for these documents.

- **Cost should be logged per run from the start.** The critic agent cost ($40/night) was only
  discovered after the fact. Build cost logging into the pipeline from the first version —
  `out/api_usage_log.jsonl` records every call with timestamp, token counts, and estimated cost.

- **Per-bank calibration is a feature.** The three banks have legitimately different PDF structures.
  The correct response is explicit per-bank configuration, not a fragile general heuristic that
  semi-works on all three.

- **Merge conditions must be strict.** `continued_from_previous=True` alone is insufficient to
  identify a true row continuation. A new date-period table (e.g. Jun 2025 vs Dec 2025) can have
  the same column count and `continued_from_previous=True` while being a genuinely separate table.
  The additional condition `t.title.strip() == ""` (no new title) is required to avoid incorrectly
  stitching distinct tables.

---

## 6. Open Items & Next Steps

### 6.1 Known Remaining Issues
- **OCR spacing artefacts:** Numbers like `4 2,911` (should be `42,911`) are handled by the
  `coerce()` function which strips commas and parses numerically, but true OCR splits within a
  digit sequence are not caught. The `_HIER` prompt instructs Gemini: "If a single number is split
  by a stray render space ('2 64,680'), join it ('264,680')." Edge-case false-positive rate
  not yet formally measured.
- **Fallback trigger is structural, not numeric:** The `_reasonable()` check triggers on missing
  tables/columns/rows, not on suspiciously low row counts. A table with the correct structure
  but half the rows missing will not trigger a retry. A row-count baseline could be added using
  the TOC's expected table count as a reference.
- **Section end_page over-extension:** Sections with no explicit heading in the body (anchored
  only via token_map) can inherit `end_page` from the next section's `start_page - 1`, which
  may include trailing narrative pages. The heading-scan refinement mitigates this but does not
  eliminate it for sections whose number appears in table data (e.g. row numbers that match
  section numbers).

### 6.2 Not Yet Attempted
- Cross-quarter comparison: extracting the same table from Q4 2025 and Q1 2026 and aligning
  rows for delta analysis
- New document types: CEO/CFO presentations, earnings call transcripts
- Generalisation classifier: auto-detecting document type without manual bank configuration
- Held-out benchmark: currently only 1 hand-verified table (DBS SGD LCR p12); need ≥5 tables
  across ≥2 banks for a meaningful accuracy metric

### 6.3 Questions for Boss / Team
*(To be filled in after feedback sessions)*

---

## For Claude Code

The actual codebase is in `/Users/Qianyunhan/Desktop/testmay21/`. Key files:
- `build_toc.py` — Pass 1: TOC extraction (deterministic, zero API)
- `extract_to_excel.py` — Pass 2: Gemini extraction + Excel writer
- `CLAUDE.md` — persistent project context and guardrails for Claude Code sessions
- `RUN_GUIDE.md` — step-by-step run instructions

Entry points (run these, not the engine modules):
```bash
python build_toc.py "DBS_4Q25_Pillar3.pdf" --out out/dbs_toc.json
python extract_to_excel.py "DBS_4Q25_Pillar3.pdf" --toc out/dbs_toc.json --out out/dbs.xlsx --no-pause
```

---

## 7. Slide Deck Parsing — Development Log

CFO presentation slide decks (DBS, OCBC, UOB 4Q25) are a fundamentally different document
type from Pillar 3 disclosures. This section records the design decisions and bugs discovered
while building `DELIVERABLE/SLIDE_Extract.py`.

### 7.1 Why a Separate Script

| Dimension | Pillar 3 | CFO Slide Deck |
|---|---|---|
| Structure | TOC with section hierarchy | No TOC — each slide is independent |
| Input to Gemini | Native PDF slice (text layer) | PNG image (charts have no text layer) |
| Page count | ~100 pages | 20-30 slides |
| Content | Dense regulatory tables | Mix of charts, waterfall bridges, P&L tables, bullets |
| Chart values | Not applicable | Must read from bar heights, colours, legends |
| PASS1 applicability | Required (TOC → page ranges) | Not applicable — no TOC |

PASS1 cannot be applied to slides. PASS2's Gemini call + schema + writer pattern is reused,
but unit grouping changes: each slide is its own unit (one call per page).

### 7.2 Architecture Decision — Single Pass, Two Stages

Early consideration was a two-pass architecture:
- Pass 1: classify chart type on each slide → build `slide_directory.json`
- Pass 2: use directory to inject slide-specific prompt → extract values

**Rejected** — Pass 2 quality is entirely dependent on Pass 1 correctness. If Pass 1
misidentifies a waterfall as a trend chart, Pass 2 uses the wrong column structure.
No recovery path.

**Adopted** — single pass, two stages within one Gemini call:
1. Gemini identifies graphic elements on the slide and matches against a prompt "guide"
2. Gemini executes extraction using the matched template — classification and extraction
   share full visual context in one call.

The "guide" embedded in the prompt (patterns A–E) is the equivalent of a TOC for slides:
pre-defined column structures that Gemini selects from based on what it sees.

### 7.3 Input: Image Not PDF

Slide decks sent as PNG (3× scale) not native PDF, because:
- Chart values (bar heights, waterfall components) are purely visual — no text layer
- Slide layout is visual (coloured boxes, overlapping text, legends)
- `media_resolution=MEDIA_RESOLUTION_HIGH` set on all image parts (see E-10)

### 7.4 Schema Additions vs Pillar 3

Same base schema (`GColumn`, `GRow`, `list[str]` values). Two additions:
- `GTable.source_type` — `"table"` (printed text, high confidence) vs `"chart"` (visual,
  verify against source). Chart cells highlighted yellow in Excel output.
- `SlideExtraction.slide_title` — main slide heading, used for tab naming.

### 7.5 Known Bugs Identified (2026-06-04)

**Bug 1 — Waterfall format ambiguity in prompt**
The prompt defines Pattern B as `[S$m] [% change]` (two columns) but also says "one row per
bridge segment". Gemini is confused — sometimes transposes bridge components into columns
(Pattern B misapplied as column headers), sometimes produces diagonal sparse output where
values appear in wrong column positions. Root cause: the prompt does not give a concrete
row-by-row example of a waterfall. Fix: replace abstract Pattern B with an explicit example
showing starting value → each component row (signed) → ending value, with exact column
alignment.

**Bug 2 — Sign wrong on "Tax and others" waterfall component**
"Tax and others" (69) shows as positive in output but is a tax drag (negative) on the slide.
The SIGN RULE in the prompt instructs Gemini to read bar colour from the legend, but Gemini
is still misreading this bar's colour. Likely because "Tax and others" bar is small and its
colour is ambiguous at 3× scale. Fix candidates: (1) increase scale to 4× for waterfall
slides; (2) add explicit instruction that tax/expense components in a profit bridge are
almost always negative unless the legend specifically shows them as green.

**Bug 3 — Duplicate sheet creation on re-run**
Running the same slide multiple times creates `Slide 05 (2)`, `Slide 05 (3)` tabs instead of
overwriting. The `tab_name()` deduplication logic appends `(N)` suffix when a name is taken,
but the existing tab is not first removed before re-creating. The `--force` flag deletes the
audit `parsed.json` to force a new API call, but the workbook tab removal logic only triggers
when the tab name exactly matches — name variants with `(2)` suffix slip through. Fix: before
writing a slide tab, remove ALL existing tabs whose name starts with the slide number prefix.

**Bug 4 — coerce() eating + signs on waterfall components**
Waterfall bridge components like `+383` (positive fee income impact) are coerced by
`coerce()` to `383` (plain integer). The + sign is lost. This matters for waterfall
arithmetic verification. Fix: in `coerce()`, detect strings starting with `+` and preserve
the sign, or keep waterfall values as strings rather than coercing to numeric.

### 7.6 Why Claude's One-Shot Extraction Worked (and Why the Pipeline Struggles)

> *Written 2026-06-04 — for future reference when designing prompt architecture.*

Claude's manual one-shot extraction of the DBS waterfall slide produced correct output
without any output format constraint. The reason it worked is not because Claude is "smarter"
— it's because the task was structured differently:

**Sequential mental steps, not concurrent ones.**
Claude read the legend ("negative" printed next to the red square), understood it, then
applied it. It read "Net profit" and recognised it as a total from financial document
experience. Only after understanding the content did it decide on output format — choosing
long format for the waterfall because it made downstream sense. Classification and extraction
were separate sequential acts.

**No schema pressure during visual reading.**
There were no output format rules to comply with while simultaneously reading the chart.
The output format was chosen *after* understanding the content, not constrained *during* it.

**Gemini in the pipeline is doing both simultaneously** — reading the visual and conforming
to a schema — under token pressure, with a prompt that mixes instructions for five different
chart types. That's why it breaks:
- It tries to match chart type AND extract values AND apply sign rules AND fit a column
  pattern all at once
- Pattern B (waterfall columns) gets misapplied because the column structure decision is
  made at the same moment as the value reading, before the chart is fully understood
- Sign errors occur because colour-reading and schema-fitting happen concurrently, and
  schema pressure wins over careful visual reading

**The architectural implication:**
A two-stage prompt that separates "understand the slide" from "extract the data" should
produce more reliable output than a single concurrent prompt. Stage 1: describe what you
see in plain English (chart type, legend, components). Stage 2: given your description,
now extract into the schema. This mirrors how Claude approached it manually.

This is the planned fix for Bug 1 (waterfall format ambiguity) — see §7.6 below.

---

### 7.6 Planned Fixes (not yet implemented)

1. Replace abstract Pattern B with a concrete waterfall example in the prompt
2. Add `+` sign preservation to `coerce()` for waterfall components
3. Fix duplicate tab logic: strip all `(N)` variants before writing
4. Investigate scale 4× vs 3× for better colour detection on small bars
5. Consider a pre-extraction stage where Gemini first describes the slide in plain English
   before structured extraction — gives it a "think aloud" step to correctly identify
   chart types and colour legends before committing to column structure

### 7.7 Chart Types Observed Across DBS/OCBC/UOB Decks

| Chart Type | Slides (DBS) | Expected Columns |
|---|---|---|
| Income statement + YoY | 3, 13, 14, 16 | [S$m current] [S$m prior] [YoY%] |
| Income statement + QoQ | 4, 5 | [S$m current] [QoQ%] |
| Waterfall/bridge | 3, 4, 5 | [Component→row] [S$m] [% change] |
| Trend bar (5 quarters) | 6, 7, 8, 9, 11, 12 | [1Q25] [2Q25] [3Q25] [4Q25] [1Q26] |
| Composition (stacked/donut) | 6, 7, 10 | [S$b or S$m] [% of total] |
| NPA reconciliation table | 17, 26 | multi-period: [FY24] [FY25] [1Q–4Q] |
| Capital ratios | 20 | monthly progression [Dec–Dec] |
| Dividend history | 21 | [Ordinary ¢] [Capital Return ¢] [Total ¢] |
| Commentary/bullets | 2, 15, 22, 23 | no columns — commentary rows only |
| Title/closing | 1, 30 | skip — no data |
