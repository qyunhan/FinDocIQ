"""
test_targeted_prompt.py — single-pass extraction using the targeted prompt.

Values are plain strings by default; only merged cells use the object form
{"value":"X","span":N,"merge_type":"aggregate|blank"}. This keeps output tokens
lean while preserving span encoding where it matters.

No max_output_tokens cap. No Pydantic schema enforcement — freeform JSON so
Gemini can mix str and dict in the values list without being forced to wrap
every cell in a GCell object.

Usage:
    cd /Users/Qianyunhan/Desktop/FinancialParser
    export GEMINI_API_KEY=...
    /Library/Frameworks/Python.framework/Versions/3.10/bin/python3 \
        Merge_Experiment/test_targeted_prompt.py \
        --pdf "OCBC_4Q25_Pillar 3.pdf" --pages 97 --id ocbc_p97_targeted
"""
import os, sys, json, argparse, time, re
sys.path.insert(0, os.path.dirname(__file__))

from extract_to_excel import render_images, log_usage, _page_numbers
from google import genai
from google.genai import types
from collections import Counter

MODEL = "gemini-2.5-pro"

PROMPT = """Transcribe the attached table image into JSON. Return: {"tables": [<one object per table>]}.

Each table object must have:
- title: printed table title verbatim
- label_header: header of the row-label column; "" if none
- columns: array of leaf-level column headers only — each as {"group": <group header or null>, "leaf": <column header>}. Exclude the row-label column. Exclude group headers that span sub-columns from the count.
- rows: every row top-to-bottom, each with row_id (printed number or null), row_type (data/total/section_header/sub_header/note), level (0=header or grand total, 1=primary, 2=sub-item), parent (null for level 0-1), label (verbatim), values

VALUES — keep output compact:
- Normal cell (span=1): plain string e.g. "54,485" or "-" or ""
- Merged cell (span>1): {"value": "X", "span": N, "merge_type": "aggregate"}
- Empty merged region (span>1): {"value": "", "span": N, "merge_type": "blank"}
- NEVER use objects for span=1 cells — plain strings only.
- sum of all spans in a row MUST equal number of columns.

Merged cell rule: a value is merged if it sits across multiple column slots with no vertical divider between them. Copy all values exactly as printed including commas, signs, and dashes."""

parser = argparse.ArgumentParser()
parser.add_argument("--pdf",   required=True)
parser.add_argument("--pages", type=int, nargs="+")
parser.add_argument("--id",    required=True)
parser.add_argument("--scale", type=float, default=3.0)
parser.add_argument("--model", default="gemini-2.5-flash",
                    help="Gemini model (e.g. gemini-2.5-flash, gemini-1.5-pro)")
args = parser.parse_args()

MODEL = args.model
OUT_DIR = os.path.join(os.path.dirname(__file__), "out", args.id)
os.makedirs(OUT_DIR, exist_ok=True)

api_key = os.environ.get("GEMINI_API_KEY", "")
if not api_key:
    sys.exit("GEMINI_API_KEY not set")

print(f"Model: {MODEL}  scale: {args.scale}×")
client = genai.Client(api_key=api_key)

# Freeform JSON config — no schema enforcement, no token cap
def build_freeform_json_config():
    try:
        return types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.0,
            max_output_tokens=131072,
            thinking_config=types.ThinkingConfig(thinking_budget=0)
        )
    except Exception:
        return types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.0,
            max_output_tokens=131072,
        )

config = build_freeform_json_config()
images = render_images(args.pdf, args.pages, scale=args.scale)
parts  = [types.Part.from_bytes(data=img, mime_type="image/png",
           media_resolution=types.PartMediaResolutionLevel.MEDIA_RESOLUTION_HIGH) for img in images]

with open(f"{OUT_DIR}/prompt.txt", "w") as f: f.write(PROMPT)

print(f"PDF: {args.pdf}  pages: {args.pages}")
print("Running single-pass targeted prompt (mixed str/dict values)...")

resp = None
for attempt in range(5):
    try:
        resp = client.models.generate_content(
            model=MODEL,
            contents=parts + [types.Part.from_text(text=PROMPT)],
            config=config
        )
        break
    except Exception as e:
        if "503" in str(e) or "UNAVAILABLE" in str(e):
            wait = 15 * (attempt + 1)
            print(f"  503 — retrying in {wait}s..."); time.sleep(wait)
        else:
            raise
if resp is None:
    sys.exit("Failed after 5 retries")

usage = log_usage(resp, args.id, image_used=True)
print(f"  ${usage['est_cost_usd']:.4f} ({usage['prompt_tokens']}in/{usage['output_tokens']}out)")

raw = resp.text or ""
with open(f"{OUT_DIR}/response.txt", "w") as f: f.write(raw)

# Parse freeform JSON
try:
    data = json.loads(raw.strip().lstrip("```json").lstrip("```").rstrip("```"))
except Exception as e:
    print(f"  ❌ JSON parse error: {e}")
    print(f"  Response tail: {raw[-200:]}")
    sys.exit(1)

with open(f"{OUT_DIR}/parsed.json", "w") as f: json.dump(data, f, indent=2)

if isinstance(data, list):
    tables = data
elif isinstance(data, dict):
    tables = data.get("tables", [data] if "rows" in data else [])
print(f"Extracted {len(tables)} table(s)\n")

# --- Check 1: span invariant ---
print("=== Span invariant ===")
span_issues = []
for t in tables:
    ncols = len(t.get("columns", []))
    for row in t.get("rows", []):
        vals = row.get("values", [])
        if not vals:
            continue
        total = sum(v.get("span", 1) if isinstance(v, dict) else 1 for v in vals)
        if total != ncols:
            span_issues.append(
                f"  [{t.get('title','')[:35]}] row {row.get('row_id')} "
                f"'{row.get('label','')[:30]}': sum(spans)={total} != ncols={ncols}"
            )
print("PASS — no violations" if not span_issues else "\n".join(span_issues[:20]))
if len(span_issues) > 20:
    print(f"  ... and {len(span_issues)-20} more")
print()

# --- Check 2: number multiset vs PDF text layer ---
print("=== Number multiset vs PDF text layer ===")
pdf_counts = _page_numbers(args.pdf, args.pages)
json_counts: Counter = Counter()
for t in tables:
    for row in t.get("rows", []):
        for v in row.get("values", []):
            val = v.get("value", "") if isinstance(v, dict) else v
            cleaned = re.sub(r'[,()\s%]', '', str(val))
            if cleaned and any(c.isdigit() for c in cleaned):
                json_counts[cleaned] += 1

num_issues = []
for num, cnt in json_counts.items():
    if pdf_counts[num] < cnt:
        num_issues.append(f"  '{num}' in JSON {cnt}x but PDF {pdf_counts[num]}x")
for num, cnt in pdf_counts.items():
    if len(num) <= 2: continue
    if json_counts[num] < cnt:
        num_issues.append(f"  '{num}' in PDF {cnt}x but JSON {json_counts[num]}x")
print("PASS" if not num_issues else f"{len(num_issues)} mismatches:")
for s in num_issues[:10]: print(s)
print()

# --- Check 3: merged cells found ---
print("=== All merged cells (span > 1) ===")
found = False
for t in tables:
    ncols = len(t.get("columns", []))
    for row in t.get("rows", []):
        for v in row.get("values", []):
            if isinstance(v, dict) and v.get("span", 1) > 1:
                found = True
                print(f"  [{t.get('title','')[:30]}] row {row.get('row_id')} "
                      f"'{row.get('label','')[:35]}': "
                      f"value={v['value']!r} span={v['span']}/{ncols} "
                      f"merge_type={v.get('merge_type')!r}")
if not found:
    print("  None found")

print(f"\nAudit saved to {OUT_DIR}/")
