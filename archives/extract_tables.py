"""
extract_tables.py — Phase 3: per-table Gemini extraction.

Reads:  out/step2_table_map.csv   (manifest from phase1)
        out/step2_docling/<tid>.json  (column grids from phase2)
Writes: out/step3_extracted/<tid>.json  (FaithfulTable JSON per table)

Usage:
  python3 extract_tables.py DBS_4Q25_Pillar3.pdf
  python3 extract_tables.py DBS_4Q25_Pillar3.pdf --table t038_p42
  python3 extract_tables.py DBS_4Q25_Pillar3.pdf --no-critic
"""
from __future__ import annotations
import os, sys, csv, json, time, random, datetime, threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import pypdfium2 as pdfium
import pdfplumber
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

MODEL       = "gemini-3.5-flash"
DOCLING_DIR = "out/step2_docling"
OUT_DIR     = "out/step3_extracted"
MANIFEST    = "out/step2_table_map.csv"
IMAGE_SCALE  = 2.0   # PDF points → pixels; 2.0 is sufficient for Gemini vision
CROP_PAD     = 8     # px padding around docling bbox crop
MAX_WORKERS  = 8     # parallel Gemini calls; Flash handles this easily

# --- API usage logging (experiment cost tracking) ---------------------------
USAGE_LOG_PATH     = "out/api_usage_log.jsonl"  # one JSON record appended per successful call
INPUT_PRICE_PER_M  = 1.50    # Gemini 3.5 Flash $/1M input tokens
OUTPUT_PRICE_PER_M = 9.00    # Gemini 3.5 Flash $/1M output tokens (thinking billed here too)
_usage_lock = threading.Lock()
_run_usage  = {"calls": 0, "prompt": 0, "output": 0, "thinking": 0, "cached": 0, "cost": 0.0}

# =============================================================================
# SCHEMA  (the contract between extraction and rendering)
# =============================================================================

class ColumnDef(BaseModel):
    col_id:   str
    group:    str | None
    leaf:     str
    col_span: int = Field(default=1)
    row_span: int = Field(default=1)

class RowDef(BaseModel):
    row_id:          str
    hierarchy_level: int
    parent_row_id:   str | None
    label:           str
    row_type:        str   # section_header | data | total | sub_header | note
    footnote_marks:  list[str]
    row_span:        int = Field(default=1)
    cells:           dict[str, str | float | int | None]
    cell_spans:      dict[str, int] = Field(default_factory=dict)  # col_id → col_span for merged data cells

class SelfCheck(BaseModel):
    n_rows:            int
    n_cols:            int
    totals_reconcile:  str   # true | false | not_applicable
    notes:             str | None

class FaithfulTable(BaseModel):
    table_id:       str
    reporting_date: str | None
    columns:        list[ColumnDef]
    rows:           list[RowDef]
    self_check:     SelfCheck



# =============================================================================
# PROMPTS
# =============================================================================

_SHARED = """
SOURCES:
- DOCLING CELL TEXT: [row,col] tagged values — authoritative for all numbers and text.
- DOCLING GRID HINT: col_span/row_span for merged headers — use exactly as given.
- Row hierarchy comes from row numbering and label indentation in the cell text.

ROW FIELDS:
- row_id: verbatim printed number ("1", "4a"). Synthesise "r1" only if none printed. NEVER renumber.
- hierarchy_level: 0=section header/grand total, 1=primary, 2=sub ("of which:"), 3=sub-sub.
- parent_row_id: row_id of immediate parent, null for top-level.
- row_span: rows this label merges vertically (default 1).
- row_type: section_header | data | total | note.
  Use "note" for reference/memo items (e.g. a large balance-sheet stock sitting alongside income flows
  of a completely different scale) — these are informational, not additive children. Set parent_row_id=null.

COLUMN FIELDS:
- col_id: c1, c2, c3... left-to-right.
- group: merged parent header text, or null.
- leaf: innermost column label. MUST be a string — use "" if blank, never null.
- col_span/row_span: copy exactly from DOCLING GRID HINT.

FIDELITY:
- Dash ("-", "–", "—") → literal "-", NOT zero or null.
- "#" → keep as "#".
- Parentheses = negative: "(283)" → -283.
- Strip thousands separators: 264,680 → 264680.
- Keep "%" on percentages.
- Blank cell with a visible column border → "-" (the column exists, it is just empty).
- If a column's every cell contains text identical to (or a repeat of) the row label, it is
  a label-mirror column — not a data column. Assign "-" to all its cells.
- null means the cell is PHYSICALLY ABSENT — no column line divides it — i.e. it is
  covered by a horizontal merge. Use null ONLY when the PDF shows no border, meaning
  one value visually spans multiple columns with no dividing line between them.
  Most blank cells are "-", not null. When in doubt, use "-".
- Never carry a value from the row above.
- Illegible → "ILLEGIBLE".
- Do NOT invent totals not printed.
- Capture footnote markers in footnote_marks list.
- label MUST be a string — never null. Use "" if blank.
- spans_pages: parent_row_id relationships continue across page breaks — a sub-row on page 2
  indented under a row from page 1 must reference that row's row_id as its parent_row_id.
- Return exactly ONE JSON object. If multiple tables are on the page, extract only the one requested.
- totals_reconcile MUST be the string "true", "false", or "not_applicable" — never a boolean.

Return STRICT JSON only — no prose, no markdown:
{
  "table_id": "<echo if given>",
  "reporting_date": "<31 Dec 2025 or null>",
  "columns": [{"col_id":"c1","group":null,"leaf":"<header>","col_span":1,"row_span":1}],
  "rows": [
    {"row_id":"1","hierarchy_level":1,"parent_row_id":null,"label":"<verbatim>","row_type":"data","footnote_marks":[],"row_span":1,"cells":{"c1":"<val>","c2":"-","c3":"-"}},
    {"row_id":"2","hierarchy_level":0,"parent_row_id":null,"label":"Total","row_type":"total","footnote_marks":[],"row_span":1,"cells":{"c1":null,"c2":null,"c3":"<merged_val>"}}
  ],
  "self_check": {"n_rows":0,"n_cols":0,"totals_reconcile":"true|false|not_applicable","notes":""}
}
"""

PROMPTS = {
    "single": (
        "Extract the single financial table below completely.\n" + _SHARED
    ),
    "multiple_on_page": (
        "Extract ONLY the table: title='{title}' date={dates}. Ignore all others on the page.\n"
        "TWO-PANEL NOTE: if two tables share a page, each has its own header row — extract only the named one.\n"
        + _SHARED
    ),
    "spans_pages": (
        "The cell text below comes from ONE table that spans multiple pages. Return ONE unified table.\n"
        "Do NOT restart row numbering at each page. Do NOT invent a total row until the last page's data.\n"
        + _SHARED
    ),
    "multiple_spanning": (
        "Extract ONLY the table for date={dates}. These fragments span pages due to width.\n"
        "Stitch into ONE table; row ids run continuously. Keep each fragment's columns separate.\n"
        + _SHARED
    ),
}


_CRITIC_PROMPT = """Audit this extraction for column assignment errors.

Each token in the x-anchored text is tagged @xNNN (measured left-edge pixel).
Tokens in the same column share the same x-coordinate (±5px).
If a value's x does not match its col_id's header x → move it to the correct col_id.
Only fix provably wrong assignments. Do not change labels, hierarchy, or row_type.
Return the corrected JSON in the same schema.

EXTRACTED JSON:
{extracted_json}

X-ANCHORED TEXT:
{page_text}
"""

# =============================================================================
# GEMINI HELPERS
# =============================================================================

_JSON_CONFIG = types.GenerateContentConfig(
    response_mime_type="application/json",
    temperature=0.0,
    max_output_tokens=8192,   # single table JSON never needs more; prevents runaway billing
)

def _log_usage(resp, label: str, attempt: int) -> None:
    """Append one JSONL record of token usage for a single call. Thread-safe.
    Captures thinking_tokens separately — they are billed as output but invisible
    in the response text. Never raises into the caller."""
    try:
        um        = getattr(resp, "usage_metadata", None)
        prompt_t  = getattr(um, "prompt_token_count", None) or 0
        output_t  = getattr(um, "candidates_token_count", None) or 0
        thought_t = getattr(um, "thoughts_token_count", None) or 0
        cached_t  = getattr(um, "cached_content_token_count", None) or 0
        total_t   = getattr(um, "total_token_count", None) or 0
        cost = (prompt_t / 1e6 * INPUT_PRICE_PER_M) + ((output_t + thought_t) / 1e6 * OUTPUT_PRICE_PER_M)
        rec = {
            "ts": datetime.datetime.now().isoformat(timespec="seconds"),
            "label": label, "model": MODEL, "attempt": attempt,
            "prompt_tokens": prompt_t, "output_tokens": output_t,
            "thinking_tokens": thought_t, "cached_tokens": cached_t,
            "total_tokens": total_t, "est_cost_usd": round(cost, 5),
        }
    except Exception as e:
        prompt_t = output_t = thought_t = cached_t = 0
        cost = 0.0
        rec = {"ts": datetime.datetime.now().isoformat(timespec="seconds"),
               "label": label, "model": MODEL, "attempt": attempt,
               "error": f"usage_capture_failed: {e}"}
    with _usage_lock:
        _run_usage["calls"]    += 1
        _run_usage["prompt"]   += prompt_t
        _run_usage["output"]   += output_t
        _run_usage["thinking"] += thought_t
        _run_usage["cached"]   += cached_t
        _run_usage["cost"]     += cost
        try:
            os.makedirs(os.path.dirname(USAGE_LOG_PATH) or ".", exist_ok=True)
            with open(USAGE_LOG_PATH, "a") as f:
                f.write(json.dumps(rec) + "\n")
        except Exception:
            pass


def _gemini(client, contents, label: str = "?", max_attempts=4) -> str:
    """Call Gemini with retry on 429/503. Returns raw response text.
    Logs token usage for every successful call to USAGE_LOG_PATH."""
    for attempt in range(max_attempts):
        try:
            resp = client.models.generate_content(model=MODEL, contents=contents, config=_JSON_CONFIG)
            _log_usage(resp, label, attempt)
            return resp.text.strip()
        except Exception as e:
            err = str(e)
            if ("429" in err or "503" in err) and attempt < max_attempts - 1:
                wait = (2 ** attempt) + random.uniform(0, 1)
                print(f"    ⏳ {'Rate limited' if '429' in err else 'Unavailable'}, retry in {wait:.1f}s...")
                time.sleep(wait)
            else:
                raise

def _parse_json(raw: str) -> dict | list:
    """Strip markdown fences and parse JSON. Handles concatenated objects by
    taking only the first complete JSON value (Gemini occasionally returns two
    tables when asked for one)."""
    if raw.startswith("```json"):
        raw = raw[7:-3].strip()
    elif raw.startswith("```"):
        raw = raw[3:-3].strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Take the first complete JSON object/array using the decoder's index
        decoder = json.JSONDecoder()
        obj, _ = decoder.raw_decode(raw.strip())
        return obj

def _unwrap(parsed, dates_val: str) -> dict:
    """Gemini occasionally wraps response in a list — unwrap to single dict."""
    if not isinstance(parsed, list):
        return parsed
    if len(parsed) == 1:
        return parsed[0]
    # Multiple tables returned — pick the one matching our target date
    matches = [t for t in parsed if dates_val.lower() in str(t.get("reporting_date", "")).lower()]
    if matches:
        print(f"    ℹ️  Multiple tables in response — selected '{dates_val}'")
        return matches[0]
    print(f"    ⚠️  Multiple tables, no date match — using first")
    return parsed[0]

# =============================================================================
# PDF HELPERS
# =============================================================================


def _border_cell_spans(pdf_path: str, page_nums: list[int], extracted: dict) -> dict:
    """
    Derive cell_spans AND corrected cell values for every data row by reading
    vertical borders directly from the PDF via pdfplumber.

    pdfplumber.table.extract() returns None for a cell that has no left vertical
    border — it is physically merged with the cell to its left. This is
    authoritative: no guessing, no Gemini, no docling heuristics.

    When a merged range contains the value anywhere within it (Gemini may have
    misassigned the column), we find the non-empty value inside the range and
    place it at the anchor col.

    Returns: {row_id: {"spans": {col_id: int}, "cells": {col_id: value}}}
    Only rows where at least one span > 1 are included.
    """
    columns     = extracted.get("columns", [])
    rows        = extracted.get("rows", [])
    col_ids     = [c["col_id"] for c in columns]
    n_data_cols = len(col_ids)

    by_rid = {str(r["row_id"]): r for r in rows}

    # Detect how many leading cols in the pdfplumber table are non-data
    # (row_id + label cols). Strategy: the extracted JSON tells us the row_ids;
    # find which pdfplumber col contains those row_id strings, then the data
    # cols are everything after the last non-data col.
    # We scan the first data page to determine the offset.
    data_col_offset = 2   # default: col0=row_no, col1=label
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_nums[0] - 1]
            tables = page.find_tables()
            if tables:
                sample_rows = tables[0].extract()
                # Find a row whose col 0 matches a known row_id
                known_rids = set(by_rid.keys())
                for srow in sample_rows:
                    if srow[0] is None:
                        continue
                    rid_candidate = (srow[0] or "").strip()
                    if rid_candidate in known_rids:
                        # Find the last col before the first data col:
                        # data cols have numeric or dash content and are NOT None.
                        # Non-data (row_no, label) cols are the leading non-None text cols.
                        offset = 0
                        for j, v in enumerate(srow):
                            if v is None:
                                break
                            # If this col's text matches a col_id from the extracted schema,
                            # we've gone too far — stop. Otherwise count it as non-data.
                            text = (v or "").strip()
                            # Heuristic: non-data cols are either the row_id or free text labels
                            # (long text, or matches known row_id). Data cols are short numeric/dash.
                            try:
                                float(text.replace(",", "").replace("-", "0").replace("(", "-").replace(")", ""))
                                break  # numeric = first data col
                            except ValueError:
                                offset = j + 1  # still in non-data region
                        if offset > 0:
                            data_col_offset = offset
                            break
    except Exception:
        pass  # keep default offset=2

    plumber_data_rows: list[list] = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg in page_nums:
            page   = pdf.pages[pg - 1]
            tables = page.find_tables()
            if not tables:
                continue
            for row in tables[0].extract():
                if row[0] is None:
                    continue
                first = (row[0] or "").strip()
                if not first:
                    continue
                # Keep rows whose col 0 matches a known row_id (any string format)
                if first in by_rid:
                    plumber_data_rows.append(row)

    result: dict[str, dict] = {}

    for plumber_row in plumber_data_rows:
        rid = (plumber_row[0] or "").strip()
        if rid not in by_rid:
            continue

        extracted_cells = dict(by_rid[rid].get("cells", {}))
        data_cells = plumber_row[data_col_offset:]

        spans: dict[str, int] = {}
        cells_out: dict[str, object] = {}
        ci = 0
        while ci < n_data_cols:
            cid = col_ids[ci]
            if ci >= len(data_cells) or data_cells[ci] is None:
                ci += 1
                continue

            # Count trailing Nones = cells merged into this anchor
            run = 1
            while (ci + run < n_data_cols and
                   ci + run < len(data_cells) and
                   data_cells[ci + run] is None):
                run += 1

            if run > 1:
                spans[cid] = run
                # Find the real value within this merged range from extracted JSON.
                # Gemini may have placed it in any col within [ci, ci+run).
                # Pick the first non-null, non-empty value; fall back to pdfplumber text.
                anchor_val = None
                for k in range(run):
                    v = extracted_cells.get(col_ids[ci + k])
                    if v is not None and v != "" and v != "-":
                        anchor_val = v
                        break
                if anchor_val is None:
                    # Use pdfplumber's own text, strip space artifacts
                    raw = (data_cells[ci] or "").strip().replace(" ", "").replace(",", "")
                    anchor_val = raw if raw and raw != "-" else "-"
                cells_out[cid] = anchor_val
            else:
                # Unmerged cell — keep Gemini's value (it had the right col)
                cells_out[cid] = extracted_cells.get(cid)

            ci += run

        if spans:
            result[rid] = {"spans": spans, "cells": cells_out}

    return result


def _xanchored_text(pdf_path: str, page_nums: list[int]) -> str:
    """Words tagged with measured x-coordinate for critic column verification."""
    blocks = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg in page_nums:
            words = pdf.pages[pg - 1].extract_words(use_text_flow=False, keep_blank_chars=False)
            rows: dict[int, list] = {}
            for w in words:
                rows.setdefault(round(w["top"] / 3), []).append(w)
            lines = [
                "  ".join(f"{w['text']}@x{int(w['x0'])}" for w in sorted(row, key=lambda x: x["x0"]))
                for row in (rows[k] for k in sorted(rows))
            ]
            blocks.append(f"[PAGE {pg}]\n" + "\n".join(lines))
    return "\n\n".join(blocks)


# =============================================================================
# DOCLING GRID
# =============================================================================

def _load_docling(tid: str) -> dict | None:
    path = os.path.join(DOCLING_DIR, f"{tid}.json")
    if not os.path.exists(path):
        return None
    with open(path) as f:
        return json.load(f)

def _derive_cell_spans(dg: dict, extracted_rows: list[dict], columns: list[dict]) -> list[dict]:
    """
    Inject cell_spans deterministically from the docling grid.

    A true merge = a data-row cell is physically absent from the docling grid.
    An empty/dash cell that HAS a grid entry is NOT a merge.

    Match extracted rows to docling rows by the row_id text appearing in col 0
    of the docling grid (the printed row number). Fall back to positional match
    if col 0 text doesn't align.
    """
    if not dg:
        return extracted_rows

    cells_list = dg.get("cells", [])

    # Build presence set: (docling_row, docling_col)
    present: set[tuple[int, int]] = set()
    col_spans_at: dict[tuple[int, int], int] = {}
    for c in cells_list:
        r, ci = c["row"], c["col"]
        present.add((r, ci))
        cs = c.get("col_span", 1)
        if cs > 1:
            col_spans_at[(r, ci)] = cs

    # Build docling row-number text → docling row index (from col 0 of data rows)
    header_rows: set[int] = {c["row"] for c in cells_list if c.get("is_col_header")}
    row_label_to_drow: dict[str, int] = {}
    data_rows_d: list[int] = []
    for c in cells_list:
        if c["row"] in header_rows:
            continue
        if c["col"] == 0 and c.get("text", "").strip():
            row_label_to_drow[c["text"].strip()] = c["row"]
        if c["row"] not in data_rows_d:
            data_rows_d.append(c["row"])
    data_rows_d = sorted(set(data_rows_d))

    # col_offset = first docling col index that holds numeric data values.
    # Gemini puts row_id and label into dedicated fields, so its c1 = docling col (col_offset).
    # Exclude: is_row_header cells, col_span>1 spanning section headers, non-numeric text.
    def _is_numeric_cell(c: dict) -> bool:
        t = c.get("text", "").strip().replace(",", "").replace("-", "").replace("(", "").replace(")", "")
        try:
            float(t)
            return True
        except ValueError:
            return t == ""  # blank cells count too

    data_col_indices = {
        c["col"] for c in cells_list
        if (c["row"] not in header_rows
            and not c.get("is_row_header")
            and c.get("col_span", 1) == 1
            and _is_numeric_cell(c))
    }
    col_offset = min(data_col_indices) if data_col_indices else 0

    # Gemini's c1 = docling col (col_offset), c2 = col_offset+1, etc.
    col_ids = [col.get("col_id", f"c{i+1}") for i, col in enumerate(columns)]

    def _infer_spans_from_nulls(row: dict, cids: list[str]) -> dict:
        """Fallback for rows beyond docling's captured range.
        None = absent/merged cell; "-" = explicit empty (not a merge).
        Two patterns handled:
          A) value followed by Nones  → value spans rightward over the Nones
          B) leading Nones then value → value is at the end; first non-null
             anchor is the leftmost non-null before the run (if any), else
             treat the run as belonging to the next value (skip).
        """
        row = dict(row)
        cells = dict(row.get("cells", {}))
        spans: dict[str, int] = {}

        # Collect groups: each group is (start_ci, length) of consecutive Nones
        # bounded by non-None values on at least one side.
        # Simpler: scan left-to-right, when we see a non-None value, count
        # following Nones as its span. Leading Nones before first value: attach
        # them to the first non-None value found to their right.
        n = len(cids)
        consumed: set[int] = set()

        # First pass: attach trailing nulls to their left-anchor (value then Nones)
        ci = 0
        while ci < n:
            if ci in consumed:
                ci += 1
                continue
            val = cells.get(cids[ci])
            if val is not None:
                run = 1
                while ci + run < n and cells.get(cids[ci + run]) is None:
                    run += 1
                if run > 1:
                    spans[cids[ci]] = run
                    for skip in range(1, run):
                        consumed.add(ci + skip)
                        cells.pop(cids[ci + skip], None)
                ci += run
            else:
                ci += 1

        # Second pass: leading Nones that were not consumed — find the first
        # non-None to their right and extend its span leftward by moving the
        # anchor to the leftmost None position.
        ci = 0
        while ci < n:
            if cells.get(cids[ci]) is None and cids[ci] in cells:
                # Find right anchor
                right = ci + 1
                while right < n and cells.get(cids[right]) is None:
                    right += 1
                if right < n and cells.get(cids[right]) is not None:
                    # Move value to leftmost position
                    cells[cids[ci]] = cells.pop(cids[right])
                    existing_span = spans.pop(cids[right], 1)
                    spans[cids[ci]] = (right - ci) + existing_span - 1
                    for skip in range(1, spans[cids[ci]]):
                        if ci + skip < n:
                            cells.pop(cids[ci + skip], None)
                    ci += spans[cids[ci]]
                else:
                    ci += 1
            else:
                ci += 1

        row["cell_spans"] = spans
        row["cells"] = cells
        return row

    # Build a lookup: (docling_row, docling_col) → cell text
    docling_values: dict[tuple[int,int], str] = {
        (c["row"], c["col"]): c.get("text", "").strip()
        for c in cells_list
        if not c.get("is_col_header") and not c.get("is_row_header")
    }

    def _apply_spans(row: dict, d_row: int) -> dict:
        row = dict(row)
        # Re-map values from docling ground truth to correct col positions.
        # Gemini may pack values left-to-right ignoring absent cols; docling knows
        # which cols actually had content (present set) and what their values were.
        cells: dict = {}
        spans: dict[str, int] = {}
        ci = 0
        while ci < len(col_ids):
            cid = col_ids[ci]
            d_ci = ci + col_offset

            if (d_row, d_ci) not in present:
                # Absent in docling = no border = covered by a merge from the previous col.
                # Skip — the owning cell's trailing-absent count will cover this col.
                ci += 1
                continue

            # Col is present. Get value from docling (ground truth for position).
            raw = docling_values.get((d_row, d_ci), "")
            if raw == "" or raw is None:
                val = None
            elif raw in ("-", "–", "—"):
                val = "-"
            else:
                val = raw.replace(",", "")

            cs = col_spans_at.get((d_row, d_ci), 1)
            if cs > 1:
                cells[cid] = val
                spans[cid] = cs
                ci += cs
            else:
                # Count consecutive absent cols after this present cell = implicit merge
                run = 1
                while (ci + run < len(col_ids) and
                       (d_row, ci + run + col_offset) not in present):
                    run += 1
                cells[cid] = val
                if run > 1:
                    spans[cid] = run
                ci += run
        row["cell_spans"] = spans
        row["cells"] = cells
        return row

    result = []
    positional_counter = 0
    for row in extracted_rows:
        if row.get("row_type") in ("section_header", "note"):
            result.append(row)
            continue

        # Try label match first (row_id in col 0 of docling)
        rid = str(row.get("row_id", "")).strip()
        d_row = row_label_to_drow.get(rid)

        if d_row is None:
            # Fallback: positional match against remaining docling rows
            if positional_counter < len(data_rows_d):
                d_row = data_rows_d[positional_counter]
            else:
                # Beyond docling's captured rows — infer merges from null runs.
                # None = absent/merged; "-" = explicitly empty (not a merge).
                result.append(_infer_spans_from_nulls(row, col_ids))
                positional_counter += 1
                continue

        positional_counter += 1
        result.append(_apply_spans(row, d_row))

    return result


def _docling_grid_hint(g: dict, first_data_col: int = 0, row_header_cols: set | None = None) -> str:
    """Compact grid hint — only DATA column header cells.
    Row-header cols (row numbers, label mirrors) are stripped entirely so
    Gemini builds exactly the right number of c1..cN columns.

    first_data_col: docling col index that maps to c1.
    row_header_cols: set of col indices to exclude completely.
    """
    rh_cols = row_header_cols or set()
    total_cols = g.get("n_cols", 0)
    data_n_cols = total_cols - len(rh_cols)

    lines = [f"n_rows={g['n_rows']} n_cols={data_n_cols}  (row-header cols excluded — build exactly {data_n_cols} columns: c1..c{data_n_cols})"]

    # Re-map docling col indices to c1..cN skipping row-header cols
    data_col_indices = [i for i in range(total_cols) if i not in rh_cols]
    docling_to_cid = {di: f"c{ci+1}" for ci, di in enumerate(data_col_indices)}

    for c in g.get("cells", []):
        col = c["col"]
        if col in rh_cols:
            continue  # strip row-header col entirely
        if not (c.get("is_col_header") or c.get("is_row_header")):
            continue
        cid = docling_to_cid.get(col, f"c?{col}")
        span_cols = c.get("col_span", 1)
        span_rows = c.get("row_span", 1)
        span = ""
        if span_cols > 1:
            span += f" col_span={span_cols}"
        if span_rows > 1:
            span += f" row_span={span_rows}"
        lines.append(f"  {cid}[row={c['row']}]{span} {repr(c['text'][:60])}")

    return "\n\nDOCLING GRID (authoritative for column count and col_span/row_span):\n" + "\n".join(lines)

# =============================================================================
# EXTRACTION MODES
# =============================================================================

def _single_pass(client, contents: list, dates_val: str, label: str = "?") -> dict:
    raw = _gemini(client, contents, label=label)
    parsed = _unwrap(_parse_json(raw), dates_val)
    return parsed



def _critic(client, extracted: dict, pdf_path: str, page_nums: list[int], label: str = "?") -> dict:
    text = _xanchored_text(pdf_path, page_nums)
    prompt = (_CRITIC_PROMPT
              .replace("{extracted_json}", json.dumps(extracted, indent=2))
              .replace("{page_text}", text))
    try:
        return _parse_json(_gemini(client, [prompt], label=label))
    except Exception as e:
        print(f"    ⚠️  Critic failed ({e}), keeping original")
        return extracted

# =============================================================================
# MAIN EXTRACTION LOOP
# =============================================================================

def extract_tables(pdf_path: str, target_id: str | None = None, no_critic: bool = True):
    if not os.path.exists(MANIFEST):
        sys.exit(f"{MANIFEST} not found — run phase1.py first.")
    os.makedirs(OUT_DIR, exist_ok=True)

    client = genai.Client()

    with open(MANIFEST) as f:
        all_rows = list(csv.DictReader(f))

    if target_id:
        rows = [r for r in all_rows if r["table_id"] == target_id]
        if not rows:
            sys.exit(f"'{target_id}' not found in manifest.")
        print(f"🎯 Single-table mode: {target_id}")
    else:
        rows = all_rows
        print(f"🚀 Extracting {len(rows)} tables via {MODEL}...")

    # Build work list — skip already done
    todo = [r for r in rows if not os.path.exists(os.path.join(OUT_DIR, f"{r['table_id']}.json"))]
    skipped = len(rows) - len(todo)
    if skipped:
        print(f"  ⏭️  Skipping {skipped} already-done tables")

    # Pre-run cost estimate (rough: ~8K input tokens/page, 4K output tokens/table).
    # NOTE: this ignores thinking tokens — actual cost is logged per-call to
    # USAGE_LOG_PATH and summarised at the end of the run.
    INPUT_TOK_PER_PAGE  = 8_000
    OUTPUT_TOK_PER_TABLE = 4_000
    est_pages = sum(int(r.get("n_pages", 1)) for r in todo)
    est_input  = est_pages * INPUT_TOK_PER_PAGE
    est_output = len(todo) * OUTPUT_TOK_PER_TABLE
    est_cost   = (est_input / 1e6 * INPUT_PRICE_PER_M) + (est_output / 1e6 * OUTPUT_PRICE_PER_M)
    critic_note = " × 2 (critic ON)" if not no_critic else ""
    print(f"  💰 Estimate: ~{est_input//1000}K input + ~{est_output//1000}K output tokens{critic_note}"
          f" ≈ ${est_cost * (2 if not no_critic else 1):.2f}")

    # Pre-extract raw page text for layouts that need it (PDFium not thread-safe)
    multi_pages = sorted({
        int(p)
        for r in todo
        if r.get("layout") in ("multiple_on_page", "spans_pages", "multiple_spanning")
        for p in r["pages"].split("+")
    })
    page_text_cache: dict[int, str] = {}
    if multi_pages:
        print(f"  📄 Pre-extracting text for {len(multi_pages)} multiple_on_page pages...")
        pdf = pdfium.PdfDocument(pdf_path)
        for pg in multi_pages:
            page_text_cache[pg] = pdf[pg - 1].get_textpage().get_text_range()

    def _process(row: dict) -> str:
        tid     = row["table_id"]
        layout  = row.get("layout", "single")
        out_file = os.path.join(OUT_DIR, f"{tid}.json")

        page_nums   = [int(p) for p in row["pages"].split("+")]
        table_title = row.get("table_title", "")
        dates_val   = row.get("dates", "") or (table_title.rsplit(" - ", 1)[-1].strip() if " - " in table_title else "")

        docling  = _load_docling(tid)
        dg       = (docling or {}).get("docling_grid") or {}
        n_cols_d = dg.get("n_cols", 0)

        # multiple_on_page: docling detects one grid per page even when two tables
        # share it — both siblings get identical cell text. Use full page text so
        # Gemini can use the title/date discriminator to pick the right panel.
        #
        # spans_pages: docling TableFormer truncates long tables that span pages,
        # missing rows on later pages. Use full page text to get all rows.
        #
        # single/other with good grid: docling cell text is isolated and complete.
        row_header_cols: set[int] = set()  # populated in docling branch, used by grid hint
        if layout in ("multiple_on_page", "spans_pages", "multiple_spanning"):
            raw = "\n\n".join(f"[PAGE {pg}]\n{page_text_cache[pg]}" for pg in page_nums)
            label = "full page — extract only the panel matching title/date above" if layout == "multiple_on_page" else "full pages"
            text_block = f"\n\nRAW PAGE TEXT ({label}):\n" + raw
        elif n_cols_d >= 2:
            # Determine which docling cols are row-header cols (row numbers, labels)
            # so we don't feed them as data cells — Gemini already gets label/row_id
            # from dedicated fields and must not treat them as c1/c2 values.
            all_cells = dg.get("cells", [])
            row_header_cols: set[int] = {
                c["col"] for c in all_cells if c.get("is_row_header")
            }
            # Also exclude col 0 if every non-header cell is a plain integer (row numbers like 1, 2, 3)
            col0_vals = [c["text"].strip() for c in all_cells
                         if c["col"] == 0 and not c.get("is_col_header") and c.get("text","").strip()]
            if col0_vals and all(v.isdigit() for v in col0_vals):
                row_header_cols.add(0)

            # Detect label-repeat columns: a non-header col whose values are all
            # long text (>6 chars, non-numeric) — these are label mirrors that
            # docling didn't mark is_row_header but should not become c1/c2.
            n_data_rows = dg.get("n_rows", 0)
            for col_idx in range(dg.get("n_cols", 0)):
                if col_idx in row_header_cols:
                    continue
                col_vals = [c["text"].strip() for c in all_cells
                            if c["col"] == col_idx
                            and not c.get("is_col_header")
                            and c.get("text", "").strip()]
                if not col_vals:
                    continue
                def _is_numeric(t):
                    return t.replace(",","").replace(".","").replace("-","").replace("(","").replace(")","").replace("%","").isdigit()
                n_text = sum(1 for v in col_vals if not _is_numeric(v) and v not in ("-", ""))
                # If the majority of cells are non-numeric text it's a label col
                if n_text >= max(1, len(col_vals) * 0.6):
                    row_header_cols.add(col_idx)

            cell_text = "\n".join(
                f"[{c['row']},{c['col']}] {c['text']}"
                for c in all_cells
                if c.get("text", "").strip()
                and not c.get("is_row_header")
                and c["col"] not in row_header_cols
            )
            text_block = "\n\nDOCLING CELL TEXT (row,col → value):\n" + cell_text
        else:
            pdf = pdfium.PdfDocument(pdf_path)
            raw = "\n\n".join(f"[PAGE {pg}]\n{pdf[pg-1].get_textpage().get_text_range()}" for pg in page_nums)
            text_block = "\n\nRAW PAGE TEXT:\n" + raw

        base_prompt = PROMPTS.get(layout, PROMPTS["single"])
        if layout in ("multiple_on_page", "multiple_spanning"):
            base_prompt = base_prompt.replace("{title}", table_title).replace("{dates}", dates_val)
        prompt = base_prompt + text_block
        if n_cols_d >= 2:
            first_data_col = min(row_header_cols) + len(row_header_cols) if row_header_cols else 0
            prompt += _docling_grid_hint(dg, first_data_col=first_data_col, row_header_cols=row_header_cols)

        parsed = _single_pass(client, [prompt], dates_val, label=tid)
        parsed["table_id"] = tid
        if not parsed.get("reporting_date"):
            parsed["reporting_date"] = dates_val

        # Inject cell_spans deterministically from docling — absent cells = true merges.
        # Skip multiple_on_page (docling grid is shared across both sibling tables).
        if layout not in ("multiple_on_page", "multiple_spanning") and n_cols_d >= 2:
            parsed["rows"] = _derive_cell_spans(dg, parsed.get("rows", []), parsed.get("columns", []))

        if not no_critic:
            parsed = _critic(client, parsed, pdf_path, page_nums, label=f"{tid}:critic")

        try:
            out_text = FaithfulTable(**parsed).model_dump_json(indent=2)
            valid = True
        except Exception as ve:
            out_text = json.dumps(parsed, indent=2)
            valid = False

        with open(out_file, "w") as f:
            f.write(out_text)

        sc = parsed.get("self_check", {})
        tag = "✅" if valid else "⚠️ "
        return f"  {tag} {tid} ({layout}, {sc.get('n_rows','?')}r {sc.get('n_cols','?')}c)"

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(_process, row): row["table_id"] for row in todo}
        for fut in as_completed(futures):
            tid = futures[fut]
            try:
                print(fut.result())
            except Exception as e:
                print(f"  ❌ FAILED {tid}: {e}")

    print(f"\n🎉 Done. Results in {OUT_DIR}/")

    u = _run_usage
    print(f"\n  📊 Actual usage this run ({u['calls']} calls):")
    print(f"     input={u['prompt']:,}  output={u['output']:,}  "
          f"thinking={u['thinking']:,}  cached={u['cached']:,} tokens")
    print(f"     measured cost ≈ ${u['cost']:.2f}  (thinking included; per-call log → {USAGE_LOG_PATH})")
    if u['output'] and u['thinking'] > u['output']:
        print(f"     ⚠️  thinking tokens ({u['thinking']:,}) exceed visible output "
              f"({u['output']:,}) — disable/cap thinking_config to cut cost")

# =============================================================================
# CLI
# =============================================================================

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("Usage: python3 extract_tables.py <pdf> [--table <id>] [--no-critic]")
    pdf = sys.argv[1]
    target = sys.argv[sys.argv.index("--table") + 1] if "--table" in sys.argv else None
    # Critic is OFF by default — adds a second Gemini call per table (doubles cost).
    # Enable explicitly with --critic only when needed.
    use_critic = "--critic" in sys.argv and "--no-critic" not in sys.argv
    extract_tables(pdf, target_id=target, no_critic=not use_critic)
