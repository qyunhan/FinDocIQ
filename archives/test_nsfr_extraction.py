"""
Test script: extract an NSFR section from any bank PDF with the updated schema.

Usage:
    GEMINI_API_KEY=... python3 test_nsfr_extraction.py                         # OCBC default
    GEMINI_API_KEY=... python3 test_nsfr_extraction.py --pdf DBS_4Q25_Pillar3.pdf --pages 84 85 86 87 --id dbs_nsfr
    GEMINI_API_KEY=... python3 test_nsfr_extraction.py --pdf OCBC_4Q25_Pillar\ 3.pdf --pages 97 98 --id ocbc_nsfr

Checks:
  1. Span invariant holds for all rows
  2. Number multiset matches PDF text layer
  3. No span > ncols (over-merge detection)
  4. All merged cells printed for manual review
"""
import os, sys, json, argparse, time
sys.path.insert(0, ".")

from extract_to_excel import (
    Extraction, GCell, build_config, validate_spans, validate_numbers,
    cut_pdf, _to_extraction, log_usage, build_prompt
)
from google import genai
from google.genai import types

MODEL = "gemini-2.5-flash"

parser = argparse.ArgumentParser()
parser.add_argument("--pdf",   default="OCBC_4Q25_Pillar 3.pdf")
parser.add_argument("--pages", type=int, nargs="+", default=[97, 98])
parser.add_argument("--id",    default="nsfr_test")
args = parser.parse_args()

PDF     = args.pdf
PAGES   = args.pages
OUT_DIR = f"out/audit/{args.id}"

os.makedirs(OUT_DIR, exist_ok=True)

api_key = os.environ.get("GEMINI_API_KEY", "")
if not api_key:
    sys.exit("GEMINI_API_KEY not set")

client = genai.Client(api_key=api_key)

unit = {
    "unit_id": args.id,
    "pages":   PAGES,
    "type":    "spanning",
    "pdf_path": PDF,
    "leaves":  [{"title": "Net Stable Funding Ratio", "number": ""}],
}

prompt   = build_prompt(unit)
pdf_bytes = cut_pdf(PDF, PAGES)

with open(f"{OUT_DIR}/pages.pdf", "wb") as f: f.write(pdf_bytes)
with open(f"{OUT_DIR}/prompt.txt", "w") as f: f.write(prompt)

print(f"PDF: {PDF}  pages: {PAGES}")
print("Calling Gemini...")
pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
config   = build_config(with_thinking=False)
for attempt in range(5):
    try:
        resp = client.models.generate_content(
            model=MODEL, contents=[pdf_part, prompt], config=config
        )
        break
    except Exception as e:
        if "503" in str(e) or "UNAVAILABLE" in str(e):
            wait = 15 * (attempt + 1)
            print(f"  503 — retrying in {wait}s (attempt {attempt+1}/5)...")
            time.sleep(wait)
        else:
            raise

with open(f"{OUT_DIR}/response.txt", "w") as f: f.write(resp.text or "")

usage = log_usage(resp, args.id, image_used=False)
print(f"Cost: ${usage['est_cost_usd']:.5f}  "
      f"({usage['prompt_tokens']} in / {usage['output_tokens']} out)\n")

with open(f"{OUT_DIR}/meta.json", "w") as f:
    json.dump({"unit_id": args.id, "pages": PAGES, "model": MODEL, "usage": usage}, f, indent=2)

ext = _to_extraction(resp)
with open(f"{OUT_DIR}/parsed.json", "w") as f: f.write(ext.model_dump_json(indent=2))

print(f"Extracted {len(ext.tables)} table(s)\n")

# --- Check 1: span invariant ---
print("=== Span invariant ===")
issues = validate_spans(ext)
print("PASS — no violations" if not issues else "\n".join(issues))
print()

# --- Check 2: number multiset ---
print("=== Number multiset vs PDF text layer ===")
num_issues = validate_numbers(ext, PDF, PAGES)
print("PASS — no mismatches" if not num_issues else f"{len(num_issues)} mismatches:")
for s in num_issues[:10]: print(s)
print()

# --- Check 3: over-merge detection (span == ncols means entire row collapsed) ---
print("=== Over-merge check (span == ncols) ===")
over = []
for t in ext.tables:
    ncols = len(t.columns)
    for row in t.rows:
        for c in row.values:
            if c.span == ncols and ncols > 1:
                over.append(f"  [{t.title[:30]}] row {row.row_id} '{row.label[:30]}': "
                            f"span={c.span} == ncols={ncols}  value={c.value!r}")
print("PASS — no full-row merges" if not over else "\n".join(over))
print()

# --- Check 4: all merged cells for manual review ---
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

# --- Check 5: per-table summary ---
print("=== Table summary ===")
for t in ext.tables:
    ncols = len(t.columns)
    nrows = len(t.rows)
    span_violations = sum(
        1 for row in t.rows
        for c in row.values if c.span == ncols and ncols > 1
    )
    print(f"  {t.title[:40]:40}  ncols={ncols}  nrows={nrows}  over-merges={span_violations}")

print(f"\nAudit saved to {OUT_DIR}/")
