"""
Two-pass extraction test: structure pass first, then extraction pass.
Same call, same image — Gemini resolves merges before attempting full extraction.

Usage:
    GEMINI_API_KEY=... python3 test_two_pass.py --pdf DBS_4Q25_Pillar3.pdf --pages 84 --id dbs_p84
    GEMINI_API_KEY=... python3 test_two_pass.py --pdf "OCBC_4Q25_Pillar 3.pdf" --pages 97 --id ocbc_p97
"""
import os, sys, json, argparse, time
sys.path.insert(0, os.path.dirname(__file__))

from extract_to_excel import (
    build_config, validate_spans, validate_numbers,
    render_images, _to_extraction, log_usage
)
from google import genai
from google.genai import types

def build_freeform_config():
    """Pass 1 config — no schema constraint, plain text response."""
    try:
        return types.GenerateContentConfig(
            temperature=0.0,
            thinking_config=types.ThinkingConfig(thinking_budget=0)
        )
    except Exception:
        return types.GenerateContentConfig(temperature=0.0)

MODEL = "gemini-2.5-pro"

parser = argparse.ArgumentParser()
parser.add_argument("--pdf",   required=True)
parser.add_argument("--pages", type=int, nargs="+")
parser.add_argument("--id",    required=True)
parser.add_argument("--model", default="gemini-2.5-flash",
                    help="Gemini model (e.g. gemini-2.5-flash, gemini-1.5-pro)")
args = parser.parse_args()

MODEL = args.model
OUT_DIR = os.path.join(os.path.dirname(__file__), "out", args.id)
os.makedirs(OUT_DIR, exist_ok=True)

api_key = os.environ.get("GEMINI_API_KEY", "")
if not api_key:
    sys.exit("GEMINI_API_KEY not set")

print(f"Model: {MODEL}")

client = genai.Client(api_key=api_key)
config = build_config(with_thinking=False)
images = render_images(args.pdf, args.pages, scale=3.0)  # higher res for border detection
parts  = [types.Part.from_bytes(data=img, mime_type="image/png",
           media_resolution=types.PartMediaResolutionLevel.MEDIA_RESOLUTION_HIGH) for img in images]

# ===========================================================================
# PASS 1 — structure only: find merged cells
# ===========================================================================
PASS1_PROMPT = """Look at this table image carefully. I want to understand its merge structure before extracting values.

First, count the data columns: how many distinct leaf-level column headers are there at the bottom of the header row? Do NOT count the row-label column on the left or row-number column. Do NOT count group headers that span multiple sub-columns as separate columns.

Then, for each row in the table (top to bottom), tell me:
- The row number (if printed on the left)
- The row label (first few words)
- Whether any cells in that row span multiple columns — and if so, which value spans how many columns

Focus on data rows where a single value appears to cover multiple column slots with no vertical divider between them. Also note rows where a shaded or blank region covers multiple columns with no internal dividers.

Do NOT extract all values — just describe the merge structure in plain English. Be specific about which row numbers have merges and how many columns each merged cell covers."""

print(f"PDF: {args.pdf}  pages: {args.pages}")
print("Pass 1 — structure...")

with open(f"{OUT_DIR}/pass1_prompt.txt", "w") as f: f.write(PASS1_PROMPT)

resp1 = None
for attempt in range(5):
    try:
        resp1 = client.models.generate_content(
            model=MODEL,
            contents=parts + [types.Part.from_text(text=PASS1_PROMPT)],
            config=build_freeform_config()
        )
        break
    except Exception as e:
        if "503" in str(e) or "UNAVAILABLE" in str(e):
            wait = 15 * (attempt + 1)
            print(f"  503 — retrying in {wait}s..."); time.sleep(wait)
        else:
            raise
if resp1 is None:
    sys.exit("Pass 1 failed after 5 retries")

usage1 = log_usage(resp1, f"{args.id}_pass1", image_used=True)
print(f"  ${usage1['est_cost_usd']:.4f} ({usage1['prompt_tokens']}in/{usage1['output_tokens']}out)")

with open(f"{OUT_DIR}/pass1_response.txt", "w") as f: f.write(resp1.text or "")
print(f"  Merge map:\n{resp1.text}\n")

# ===========================================================================
# PASS 2 — full extraction using merge map from pass 1
# ===========================================================================
PASS2_PROMPT = f"""Now extract the full table using the merge map identified in the previous step.

Merge map from structure analysis:
{resp1.text}

Using this merge map, transcribe every table in this image into a JSON structure exactly as printed.
Each row must include row_id (printed number or null), row_type (data/total/section_header/sub_header/note),
level (0=header or grand total, 1=primary, 2=sub-item, 3=sub-sub-item), parent (row_id of nearest row
one level above; null for level 0 and 1), label (verbatim), and values.

For columns: count only leaf-level headers. Group headers spanning sub-columns are NOT separate columns.
Scope/currency labels (e.g. "Group - ALL Currency") are NOT columns.

For values: use the merge map above to set span correctly.
- Normal cell: plain string e.g. "54,485"
- Merged cell with value: {{"value":"X","span":N,"merge_type":"aggregate"}}
- Merged empty region: {{"value":"","span":N,"merge_type":"blank"}}
- Sum of all spans in a row MUST equal total number of leaf columns.

Copy every value exactly as printed. Never invent, reorder, or omit any row, column, or value.
Every section/category label must appear as a section_header row."""

print("Pass 2 — extraction...")
with open(f"{OUT_DIR}/pass2_prompt.txt", "w") as f: f.write(PASS2_PROMPT)

resp2 = None
for attempt in range(5):
    try:
        resp2 = client.models.generate_content(
            model=MODEL,
            contents=parts + [types.Part.from_text(text=PASS2_PROMPT)],
            config=config
        )
        break
    except Exception as e:
        if "503" in str(e) or "UNAVAILABLE" in str(e):
            wait = 15 * (attempt + 1)
            print(f"  503 — retrying in {wait}s..."); time.sleep(wait)
        else:
            raise
if resp2 is None:
    sys.exit("Pass 2 failed after 5 retries")

usage2 = log_usage(resp2, f"{args.id}_pass2", image_used=True)
print(f"  ${usage2['est_cost_usd']:.4f} ({usage2['prompt_tokens']}in/{usage2['output_tokens']}out)")
total_cost = usage1['est_cost_usd'] + usage2['est_cost_usd']
print(f"Total cost: ${total_cost:.4f}\n")

with open(f"{OUT_DIR}/pass2_response.txt", "w") as f: f.write(resp2.text or "")
with open(f"{OUT_DIR}/meta.json", "w") as f:
    json.dump({"pass1": usage1, "pass2": usage2, "total_cost": total_cost}, f, indent=2)

ext = _to_extraction(resp2)
with open(f"{OUT_DIR}/parsed.json", "w") as f: f.write(ext.model_dump_json(indent=2))
print(f"Extracted {len(ext.tables)} table(s)\n")

# --- Check 1: span invariant ---
print("=== Span invariant ===")
issues = validate_spans(ext)
print("PASS — no violations" if not issues else "\n".join(issues))
print()

# --- Check 2: number multiset ---
print("=== Number multiset vs PDF text layer ===")
num_issues = validate_numbers(ext, args.pdf, args.pages)
print("PASS — no mismatches" if not num_issues else f"{len(num_issues)} mismatches:")
for s in num_issues[:10]: print(s)
print()

# --- Check 3: merged cells ---
print("=== All merged cells (span > 1) ===")
found = False
for t in ext.tables:
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

# --- Check 4: table summary ---
print("=== Table summary ===")
for t in ext.tables:
    ncols  = len(t.columns)
    merges = sum(1 for row in t.rows for c in row.values if c.span > 1)
    over_m = sum(1 for row in t.rows for c in row.values if c.span == ncols and ncols > 1)
    print(f"  {t.title[:40]:40}  ncols={ncols}  nrows={len(t.rows)}  merges={merges}  over-merges={over_m}")

print(f"\nAudit saved to {OUT_DIR}/")
