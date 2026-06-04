"""
Test script: extract any table section from any bank PDF, one page per call.
Detects merged data cells deterministically from PDF text layer (column boundary
+ value x-position) and injects as a factual hint into the prompt.

Usage:
    GEMINI_API_KEY=... python3 test_extraction.py --pdf DBS_4Q25_Pillar3.pdf --pages 84 85 86 87 --id dbs_nsfr
    GEMINI_API_KEY=... python3 test_extraction.py --pdf "OCBC_4Q25_Pillar 3.pdf" --pages 97 98 --id ocbc_nsfr

Checks:
  1. Span invariant holds for all rows (sum of spans == ncols per row)
  2. Number multiset matches PDF text layer
  3. No full-row over-merges (span == ncols)
  4. All merged cells printed for manual review
"""
import os, sys, json, argparse, time
sys.path.insert(0, os.path.dirname(__file__))

from extract_to_excel import (
    build_config, validate_spans, validate_numbers,
    render_images, _to_extraction, log_usage, Extraction
)
from google import genai
from google.genai import types

MODEL = "gemini-2.5-pro"

# ---------------------------------------------------------------------------
# Deterministic merge detection from PDF text layer
# ---------------------------------------------------------------------------
def detect_merges_text(pdf_path: str, page_1based: int) -> list[dict]:
    """Detect merged data cells. Two strategies depending on PDF type:

    Strategy 1 — pdfplumber native table (DBS, UOB — grid line PDFs):
      pdfplumber returns None for merged cell continuations. Reliable and exact.

    Strategy 2 — column boundary + empty active column analysis (OCBC — no grid):
      Uses rect-derived column boundaries + text x-positions. A numeric value
      whose column has contiguous empty active columns to its left is merged.

    Returns list of:
      {"label": str, "start_col": int, "span": int, "value": str}
    Only data-row merges (span > 1). Returns [] if none found.
    """
    try:
        import pdfplumber
    except ImportError:
        return []

    import re
    from collections import Counter

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_1based - 1]

        # ---------------------------------------------------------------
        # Strategy 1: pdfplumber native table (grid-line PDFs)
        # ---------------------------------------------------------------
        tables = page.extract_tables()
        main_table = next(
            (t for t in sorted(tables, key=len, reverse=True)
             if len(t) >= 8 and t[0] and len(t[0]) >= 4),
            None
        )
        if main_table:
            merges = []
            for row in main_table:
                if not row:
                    continue
                row_id = (row[0] or "").strip()
                label  = (row[1] or "").strip() if len(row) > 1 else ""
                if not re.match(r'^\d+[a-z]?$', row_id):
                    continue
                data_cols = row[2:]
                i = 0
                while i < len(data_cols):
                    val = data_cols[i]
                    if val is not None and val.strip():
                        span = 1
                        for j in range(i + 1, len(data_cols)):
                            if data_cols[j] is None:
                                span += 1
                            else:
                                break
                        if span > 1:
                            merges.append({
                                "label":     f"{row_id} {label}"[:40],
                                "start_col": i,
                                "span":      span,
                                "value":     val.strip(),
                            })
                        i += span
                    else:
                        i += 1
            return merges  # even if empty — this PDF has grid lines, trust it

        # ---------------------------------------------------------------
        # Strategy 2: column boundary + empty active column (no-grid PDFs)
        # ---------------------------------------------------------------
        words = page.extract_words()
        if not words:
            return []

        page_h = float(page.height)
        page_w = float(page.width)
        label_cutoff  = page_w * 0.35
        header_cutoff = page_h * 0.30

        # column boundaries from thin rects (OCBC has these even without grid lines)
        vlines  = [r for r in page.rects if (r['x1']-r['x0']) < 2 and (r['y1']-r['y0']) < 2]
        col_xs  = sorted(set(round(r['x0'], 1) for r in vlines))
        if len(col_xs) < 2:
            return []

        # column centers = midpoints between consecutive boundaries
        last_gap    = col_xs[-1] - col_xs[-2]
        all_bnd     = col_xs + [col_xs[-1] + last_gap]
        col_centers = [(all_bnd[i] + all_bnd[i+1]) / 2 for i in range(len(all_bnd) - 1)]
        col_edges   = [(col_centers[i] + col_centers[i+1]) / 2 for i in range(len(col_centers) - 1)]

        def assign_col(cx):
            for i, edge in enumerate(col_edges):
                if cx < edge:
                    return i
            return len(col_centers) - 1

        def is_numeric(text):
            t = text.strip()
            return bool(re.match(r'^-$|^[\d,]+(\.\d+)?%?$|^\([\d,]+\)$', t))

        # group data words by row band
        data_words = [w for w in words
                      if float(w['top']) > header_cutoff
                      and float(w['x0']) > label_cutoff]
        rows_by_y = {}
        for w in data_words:
            y = round(float(w['top']) / 3) * 3
            rows_by_y.setdefault(y, []).append(w)

        row_presence = {}
        row_vals     = {}
        for y, ws in rows_by_y.items():
            row_presence[y] = set()
            row_vals[y]     = {}
            for w in ws:
                cx  = (float(w['x0']) + float(w['x1'])) / 2
                col = assign_col(cx)
                row_presence[y].add(col)
                row_vals[y][col] = w

        # active columns = present in >30% of data rows
        col_freq    = Counter(c for pres in row_presence.values() for c in pres)
        threshold   = max(1, len(rows_by_y) * 0.3)
        active_cols = {c for c, cnt in col_freq.items() if cnt >= threshold}

        label_by_y = {}
        for w in words:
            if float(w['top']) > header_cutoff and float(w['x0']) <= label_cutoff:
                y = round(float(w['top']) / 3) * 3
                label_by_y.setdefault(y, []).append(w['text'])

        merges = []
        for y in sorted(rows_by_y):
            present = row_presence[y]
            missing = active_cols - present
            if not missing:
                continue
            for col_idx, w in row_vals[y].items():
                if not is_numeric(w['text']):
                    continue
                # contiguous missing active cols immediately to the LEFT
                left = []
                for c in range(col_idx - 1, -1, -1):
                    if c in missing and c in active_cols:
                        left.insert(0, c)
                    else:
                        break
                if not left:
                    continue
                span = len(left) + 1
                label = " ".join(label_by_y.get(y, []))
                merges.append({
                    "label":     label[:40],
                    "start_col": min(left),
                    "span":      span,
                    "value":     w['text'],
                })

        return merges


def build_merge_hint(merges: list[dict]) -> str:
    if not merges:
        return ""
    lines = [
        "KNOWN MERGED DATA CELLS (detected from PDF text positions — treat as ground truth):",
        "In the values array, set span>1 for these rows only. All other rows have span=1.",
    ]
    for m in merges:
        lines.append(
            f"  - Row '{m['label']}': value={m['value']!r} "
            f"spans {m['span']} columns starting at col {m['start_col']} (0-based data column)"
        )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Base prompt
# ---------------------------------------------------------------------------
_BASE_PROMPT = """Transcribe the attached table into a JSON structure exactly as printed. Each row must include row_id (the printed number), row_type (data, total, or section_header), level (0 for headers, 1+ for items), parent (the ID of the row above it), label (verbatim text), and values. For columns, count only the bottom-most "leaf" headers; do not count group headers that span multiple columns as separate entries. Maintain absolute value fidelity: keep all commas, signs, and dashes exactly as shown.

The critical rule for data alignment is the Vertical Border Rule: two columns are merged if and only if there is no vertical line between them. If a value is centered across multiple columns without dividers, it is a merged cell. Represent a merged cell as {{"value": "X", "span": N, "merge_type": "aggregate"}}. The sum of all spans in a row must exactly equal the total number of leaf columns.

Never omit section headers or category labels, even if they are shaded. A shaded cell with vertical borders is NOT merged — output as normal cell with value "". Use {{"value": "", "span": N, "merge_type": "blank"}} for empty merged areas.

Note: span>1 applies to data row values only. Column headers always have span=1."""

# ---------------------------------------------------------------------------
# Args
# ---------------------------------------------------------------------------
parser = argparse.ArgumentParser()
parser.add_argument("--pdf",   required=True,           help="Path to PDF file")
parser.add_argument("--pages", type=int, nargs="+",     help="1-based page numbers to extract")
parser.add_argument("--id",    required=True,           help="Unique run ID (used for audit folder)")
parser.add_argument("--scale", type=float, default=2.0, help="Image render scale (default 2.0)")
args = parser.parse_args()

OUT_DIR = os.path.join(os.path.dirname(__file__), "out", args.id)
os.makedirs(OUT_DIR, exist_ok=True)

api_key = os.environ.get("GEMINI_API_KEY", "")
if not api_key:
    sys.exit("GEMINI_API_KEY not set")

client = genai.Client(api_key=api_key)
config = build_config(with_thinking=False)

all_tables = []
total_in   = 0
total_out  = 0
total_cost = 0.0

# --- one call per page ---
for page_num in args.pages:
    print(f"  page {page_num}...", end=" ", flush=True)

    # detect merges deterministically
    merges = detect_merges_text(args.pdf, page_num)
    merge_hint = build_merge_hint(merges)
    if merges:
        print(f"[{len(merges)} merge(s) detected] ", end="", flush=True)

    PROMPT = (_BASE_PROMPT + "\n\n" + merge_hint) if merge_hint else _BASE_PROMPT

    images = render_images(args.pdf, [page_num], scale=args.scale)

    resp = None
    for attempt in range(5):
        try:
            parts = [types.Part.from_bytes(data=img, mime_type="image/png") for img in images]
            parts.append(types.Part.from_text(text=PROMPT))
            resp = client.models.generate_content(
                model=MODEL,
                contents=parts,
                config=config,
            )
            break
        except Exception as e:
            if "503" in str(e) or "UNAVAILABLE" in str(e):
                wait = 15 * (attempt + 1)
                print(f"\n    503 — retrying in {wait}s...", end=" ", flush=True)
                time.sleep(wait)
            else:
                raise
    if resp is None:
        sys.exit(f"page {page_num} failed after 5 retries")

    usage = log_usage(resp, f"{args.id}_p{page_num}", image_used=False)
    total_in   += usage["prompt_tokens"]
    total_out  += usage["output_tokens"]
    total_cost += usage["est_cost_usd"]
    print(f"${usage['est_cost_usd']:.4f} ({usage['prompt_tokens']}in/{usage['output_tokens']}out)")

    page_dir = f"{OUT_DIR}/page_{page_num}"
    os.makedirs(page_dir, exist_ok=True)
    with open(f"{page_dir}/prompt.txt",      "w") as f: f.write(PROMPT)
    with open(f"{page_dir}/merge_hint.txt",  "w") as f: f.write(merge_hint or "(none)")
    with open(f"{page_dir}/response.txt",    "w") as f: f.write(resp.text or "")
    with open(f"{page_dir}/meta.json",       "w") as f: json.dump(usage, f, indent=2)

    ext = _to_extraction(resp)
    with open(f"{page_dir}/parsed.json",     "w") as f: f.write(ext.model_dump_json(indent=2))
    all_tables.extend(ext.tables)

print(f"\nTotal cost: ${total_cost:.5f}  ({total_in} in / {total_out} out tokens)")

combined = Extraction(tables=all_tables)
with open(f"{OUT_DIR}/combined.json", "w") as f: f.write(combined.model_dump_json(indent=2))
print(f"Extracted {len(all_tables)} table(s) across {len(args.pages)} pages\n")

# --- Check 1: span invariant ---
print("=== Span invariant ===")
issues = validate_spans(combined)
print("PASS — no violations" if not issues else "\n".join(issues))
print()

# --- Check 2: number multiset ---
print("=== Number multiset vs PDF text layer ===")
num_issues = validate_numbers(combined, args.pdf, args.pages)
print("PASS — no mismatches" if not num_issues else f"{len(num_issues)} mismatches:")
for s in num_issues[:10]: print(s)
print()

# --- Check 3: all merged cells for manual review ---
print("=== All merged cells (span > 1) ===")
found = False
for t in combined.tables:
    ncols = len(t.columns)
    for row in t.rows:
        for c in row.values:
            if c.span > 1:
                found = True
                print(f"  [{t.title[:30]}] row {row.row_id} '{row.label[:35]}': "
                      f"value={c.value!r} span={c.span}/{ncols} merge_type={c.merge_type!r}")
if not found:
    print("  None found")
print()

# --- Check 4: over-merge detection ---
print("=== Over-merge check (span == ncols = full row collapsed) ===")
over = []
for t in combined.tables:
    ncols = len(t.columns)
    for row in t.rows:
        for c in row.values:
            if c.span == ncols and ncols > 1:
                over.append(f"  [{t.title[:30]}] row {row.row_id} '{row.label[:30]}': "
                            f"span={c.span}=ncols  value={c.value!r}")
print("PASS — no full-row merges" if not over else "\n".join(over))
print()

# --- Check 5: detected merges vs extracted merges ---
print("=== Detected merges vs extracted ===")
all_detected = []
for page_num in args.pages:
    detected = detect_merges_text(args.pdf, page_num)
    all_detected.extend(detected)
if all_detected:
    for m in all_detected:
        print(f"  DETECTED: row '{m['label']}' value={m['value']!r} start_col={m['start_col']} span={m['span']}")
    # check each detected merge appears in extraction
    for m in all_detected:
        found_in_ext = False
        for t in combined.tables:
            for row in t.rows:
                for c in row.values:
                    if c.value == m['value'] and c.span == m['span']:
                        found_in_ext = True
        status = "OK" if found_in_ext else "MISSING in extraction"
        print(f"    → {status}")
else:
    print("  No merges detected on these pages")
print()

# --- Check 6: table summary ---
print("=== Table summary ===")
for t in combined.tables:
    ncols  = len(t.columns)
    nrows  = len(t.rows)
    merges = sum(1 for row in t.rows for c in row.values if c.span > 1)
    over_m = sum(1 for row in t.rows for c in row.values if c.span == ncols and ncols > 1)
    print(f"  {t.title[:40]:40}  ncols={ncols}  nrows={nrows}  merges={merges}  over-merges={over_m}")

print(f"\nAudit saved to {OUT_DIR}/")
