"""
test_single_vs_multipass.py — compare single-pass vs multi-pass extraction
on the same slide. Prints a side-by-side diff of extracted values.

Usage:
    python test_single_vs_multipass.py OCBC4Q25_CFO_presentation.pdf --slide 6
    python test_single_vs_multipass.py OCBC4Q25_CFO_presentation.pdf --slide 6 --all-visual
    python test_single_vs_multipass.py OCBC4Q25_CFO_presentation.pdf --slides 3,6,8

Outputs:
    test_results/slide_{N}_multipass.json   ← current pipeline output
    test_results/slide_{N}_singlepass.json  ← single-pass output
    test_results/slide_{N}_diff.txt         ← side-by-side diff
    test_results/summary.json               ← accuracy metrics across all tested slides

Does NOT touch your existing audit folder or Excel output.
Requires GEMINI_API_KEY in environment.
"""
from __future__ import annotations
import os, sys, json, io, re, argparse, time
from pathlib import Path
from typing import Any

import pypdfium2 as pdfium
from google import genai
from google.genai import types as gtypes

# ── Config ────────────────────────────────────────────────────────────────────
MODEL       = "gemini-2.5-flash"
IMAGE_SCALE = 3.0
OUT_DIR     = Path("test_results")

INPUT_PRICE_PER_M  = 0.30
OUTPUT_PRICE_PER_M = 2.50

# ── Image helpers ─────────────────────────────────────────────────────────────

def render_page(pdf_path: str, page_1based: int, scale: float = IMAGE_SCALE) -> bytes:
    src = pdfium.PdfDocument(pdf_path)
    pil = src[page_1based - 1].render(scale=scale).to_pil()
    buf = io.BytesIO()
    pil.save(buf, format="PNG")
    return buf.getvalue()

def img_part(img_bytes: bytes) -> gtypes.Part:
    return gtypes.Part.from_bytes(
        data=img_bytes,
        mime_type="image/png",
        media_resolution=gtypes.PartMediaResolutionLevel.MEDIA_RESOLUTION_HIGH,
    )

def call_gemini(client, parts: list, text_only: bool = False) -> tuple[str, float]:
    kwargs: dict = {"temperature": 0.0}
    if not text_only:
        kwargs["response_mime_type"] = "application/json"
    try:
        kwargs["thinking_config"] = gtypes.ThinkingConfig(thinking_budget=0)
    except Exception:
        pass
    config = gtypes.GenerateContentConfig(**kwargs)

    for attempt in range(3):
        try:
            resp = client.models.generate_content(
                model=MODEL, contents=parts, config=config
            )
            break
        except Exception as e:
            if any(c in str(e) for c in ("503", "429", "UNAVAILABLE")):
                wait = 15 * (2 ** attempt)
                print(f"  ⏳ retry in {wait}s...")
                time.sleep(wait)
            else:
                raise
    else:
        raise RuntimeError("Gemini failed after 3 retries")

    text = (resp.text or "").strip()
    um   = getattr(resp, "usage_metadata", None)
    pt   = getattr(um, "prompt_token_count", 0) or 0
    ot   = getattr(um, "candidates_token_count", 0) or 0
    cost = (pt / 1e6 * INPUT_PRICE_PER_M) + (ot / 1e6 * OUTPUT_PRICE_PER_M)
    return text, cost

def strip_fences(raw: str) -> str:
    s = raw.strip()
    if s.startswith("```"):
        s = s.split("```", 1)[1].lstrip("json").strip()
        s = s.rsplit("```", 1)[0].strip()
    return s

# ── APPROACH A: Multi-pass (current pipeline) ─────────────────────────────────

PASS1_PROMPT = """Examine this bank CFO presentation slide carefully.

═══════════════════════════════════════════════════
STEP 1 — INVENTORY
═══════════════════════════════════════════════════
List every printed number visible on the slide with its location and adjacent label.

═══════════════════════════════════════════════════
STEP 2 — STRUCTURE
═══════════════════════════════════════════════════
For each data element: type, title, structure, units.

═══════════════════════════════════════════════════
STEP 3 — VERIFY
═══════════════════════════════════════════════════
Apply arithmetic constraints. Write check explicitly. Fix any errors before proceeding.

═══════════════════════════════════════════════════
STEP 4 — PRE-MAP TO SCHEMA FIELDS
═══════════════════════════════════════════════════
For every element, map each value:
  ELEMENT {idx} | {element_type} | "{element_title}"
  series="{series}" period="{period}" value="{value}" row_type="{row_type}" level={level}

For donut dual ring: trace the callout label lines to determine which period
belongs to which ring. State this explicitly before assigning values.
series = segment name, period = period label as printed (NOT "inner"/"outer").
"""

def build_pass2_prompt(description: str) -> str:
    return f"""Convert this pre-mapped field description into JSON.
Read the Step 4 pre-mapping and transcribe it exactly. Do NOT re-interpret.

<slide_description>
{description}
</slide_description>

OUTPUT FORMAT:
{{
  "slide_title": "...",
  "elements": [
    {{
      "element_idx": 0,
      "element_type": "...",
      "element_title": "...",
      "source": "chart",
      "units": "...",
      "data_points": [
        {{
          "series": "...",
          "period": "...",
          "value": "...",
          "row_type": "data",
          "level": 1,
          "sign": null,
          "order": 0
        }}
      ]
    }}
  ]
}}

Transcribe Step 4 exactly. Return ONLY the JSON object."""

def run_multipass(client, img_bytes: bytes) -> tuple[dict, float, str]:
    """Run current 3-pass pipeline. Returns (result_dict, total_cost, description)."""
    total_cost = 0.0

    # Pass 1
    desc, c1 = call_gemini(client, [img_part(img_bytes), PASS1_PROMPT], text_only=True)
    total_cost += c1

    # Pass 2 — text only, no image
    p2_prompt = build_pass2_prompt(desc)
    raw2, c2  = call_gemini(client, [p2_prompt])
    total_cost += c2

    try:
        result = json.loads(strip_fences(raw2))
    except Exception as e:
        result = {"error": str(e), "raw": raw2[:500]}

    return result, total_cost, desc


# ── APPROACH B: Single-pass ───────────────────────────────────────────────────

SINGLE_PASS_PROMPT = """Extract ALL financial data from this bank CFO presentation slide.

Return a JSON object:
{
  "slide_title": "...",
  "elements": [
    {
      "element_idx": 0,
      "element_type": "text_table|waterfall|stacked_bar|stacked_bar_with_overlay|trend_line|kpi_grid|pie|donut_dual_ring|other",
      "element_title": "...",
      "source": "table or chart",
      "units": "S$m or % or other",
      "self_check": "arithmetic check string or null",
      "data_points": [
        {
          "series": "row/bar/segment label verbatim",
          "period": "time period or null",
          "value": "verbatim as printed",
          "row_type": "data|total|sub|start|end|bridge|note",
          "level": 1,
          "parent": null,
          "group": null,
          "sign": null,
          "order": 0
        }
      ]
    }
  ]
}

Rules:
- value: ALWAYS verbatim as printed. Never convert. "5,948" not 5948.
- Waterfall: sign="+" or "-" on every bridge component. Verify start+sum=end.
- Bold rows: row_type="total". Indented rows: level=2.
- Donut dual ring: trace the callout label lines to assign periods correctly.
- Add extra key/value pairs freely for any additional columns on the slide.
- Return ONLY the JSON object, no markdown.
"""

def run_singlepass(client, img_bytes: bytes) -> tuple[dict, float]:
    """Run single-pass extraction. Returns (result_dict, cost)."""
    raw, cost = call_gemini(client, [img_part(img_bytes), SINGLE_PASS_PROMPT])
    try:
        result = json.loads(strip_fences(raw))
    except Exception as e:
        result = {"error": str(e), "raw": raw[:500]}
    return result, cost


# ── DIFF & SCORING ────────────────────────────────────────────────────────────

def extract_values(result: dict) -> list[dict]:
    """Flatten result into list of {element, series, period, value} dicts."""
    rows = []
    for elem in result.get("elements", []):
        title = elem.get("element_title", "")
        etype = elem.get("element_type", "")
        for dp in elem.get("data_points", []):
            rows.append({
                "element": title,
                "type":    etype,
                "series":  dp.get("series", ""),
                "period":  dp.get("period", ""),
                "value":   dp.get("value", ""),
                "sign":    dp.get("sign", ""),
            })
    return rows

def normalise_value(v: str) -> str:
    """Strip commas, spaces, trailing zeros for comparison."""
    s = str(v).strip().replace(",", "").replace(" ", "")
    try:
        f = float(s.strip("()%+").replace("(", "-").replace(")", ""))
        return f"{f:.4g}"
    except Exception:
        return s.lower()

def diff_results(multi: dict, single: dict) -> dict:
    """
    Compare two extraction results.
    Returns diff report with matches, mismatches, missing rows.
    """
    mv = extract_values(multi)
    sv = extract_values(single)

    # Build lookup by (element, series, period)
    def key(r):
        return (
            r["element"].lower().strip(),
            r["series"].lower().strip(),
            str(r["period"]).lower().strip(),
        )

    multi_map  = {key(r): r for r in mv}
    single_map = {key(r): r for r in sv}

    all_keys = set(multi_map) | set(single_map)

    matches    = []
    mismatches = []
    only_multi  = []
    only_single = []

    for k in sorted(all_keys):
        mr = multi_map.get(k)
        sr = single_map.get(k)

        if mr and sr:
            mv_norm = normalise_value(mr["value"])
            sv_norm = normalise_value(sr["value"])
            if mv_norm == sv_norm:
                matches.append({
                    "key":          k,
                    "value":        mr["value"],
                    "sign_multi":   mr.get("sign", ""),
                    "sign_single":  sr.get("sign", ""),
                    "sign_match":   mr.get("sign", "") == sr.get("sign", ""),
                })
            else:
                mismatches.append({
                    "key":           k,
                    "multi_value":   mr["value"],
                    "single_value":  sr["value"],
                    "multi_sign":    mr.get("sign", ""),
                    "single_sign":   sr.get("sign", ""),
                })
        elif mr:
            only_multi.append(mr)
        else:
            only_single.append(sr)

    total       = len(all_keys)
    n_match     = len(matches)
    pct_match   = round(100 * n_match / total, 1) if total else 0

    return {
        "total_keys":    total,
        "matches":       n_match,
        "mismatches":    len(mismatches),
        "only_multi":    len(only_multi),
        "only_single":   len(only_single),
        "match_pct":     pct_match,
        "match_detail":  matches,
        "mismatch_detail": mismatches,
        "only_multi_detail":  only_multi,
        "only_single_detail": only_single,
    }

def format_diff_report(slide_num: int, diff: dict,
                        multi_cost: float, single_cost: float,
                        description: str) -> str:
    lines = []
    lines.append(f"{'='*70}")
    lines.append(f"SLIDE {slide_num:02d} — EXTRACTION COMPARISON")
    lines.append(f"{'='*70}")
    lines.append(f"Multi-pass cost:  ${multi_cost:.5f}")
    lines.append(f"Single-pass cost: ${single_cost:.5f}")
    lines.append(f"")
    lines.append(f"Total data keys:  {diff['total_keys']}")
    lines.append(f"Matches:          {diff['matches']}  ({diff['match_pct']}%)")
    lines.append(f"Mismatches:       {diff['mismatches']}")
    lines.append(f"Only in multi:    {diff['only_multi']}")
    lines.append(f"Only in single:   {diff['only_single']}")

    if diff["mismatch_detail"]:
        lines.append(f"\n{'─'*70}")
        lines.append("VALUE MISMATCHES:")
        lines.append(f"{'─'*70}")
        for m in diff["mismatch_detail"]:
            elem, series, period = m["key"]
            lines.append(f"  [{elem}] {series} / {period}")
            lines.append(f"    multi:  {m['multi_value']}  sign={m['multi_sign']!r}")
            lines.append(f"    single: {m['single_value']}  sign={m['single_sign']!r}")

    if diff["only_multi_detail"]:
        lines.append(f"\n{'─'*70}")
        lines.append("ONLY IN MULTI-PASS (missing from single):")
        for r in diff["only_multi_detail"]:
            lines.append(f"  [{r['element']}] {r['series']} / {r['period']} = {r['value']}")

    if diff["only_single_detail"]:
        lines.append(f"\n{'─'*70}")
        lines.append("ONLY IN SINGLE-PASS (missing from multi):")
        for r in diff["only_single_detail"]:
            lines.append(f"  [{r['element']}] {r['series']} / {r['period']} = {r['value']}")

    if diff["match_detail"]:
        lines.append(f"\n{'─'*70}")
        lines.append("MATCHES:")
        for m in diff["match_detail"]:
            elem, series, period = m["key"]
            sign_note = ""
            if not m["sign_match"] and (m["sign_multi"] or m["sign_single"]):
                sign_note = f"  ⚠ sign mismatch: multi={m['sign_multi']!r} single={m['sign_single']!r}"
            lines.append(f"  ✓ [{elem}] {series} / {period} = {m['value']}{sign_note}")

    lines.append(f"\n{'─'*70}")
    lines.append("PASS 1 DESCRIPTION (multi-pass):")
    lines.append(f"{'─'*70}")
    # Show Step 4 pre-mapping only
    step4_match = re.search(r"STEP 4.*?(?=\Z)", description, re.DOTALL | re.IGNORECASE)
    if step4_match:
        lines.append(step4_match.group(0)[:2000])
    else:
        lines.append(description[:1000])

    return "\n".join(lines)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def test_slide(client, pdf_path: str, slide_num: int) -> dict:
    print(f"\n{'─'*50}")
    print(f"Testing slide {slide_num:02d}...")

    img_bytes = render_page(pdf_path, slide_num)

    # Run both approaches
    print(f"  Running multi-pass...")
    multi_result, multi_cost, description = run_multipass(client, img_bytes)
    print(f"  Running single-pass...")
    single_result, single_cost = run_singlepass(client, img_bytes)

    # Diff
    diff = diff_results(multi_result, single_result)

    # Save outputs
    OUT_DIR.mkdir(exist_ok=True)
    (OUT_DIR / f"slide_{slide_num:02d}_multipass.json").write_text(
        json.dumps(multi_result, indent=2))
    (OUT_DIR / f"slide_{slide_num:02d}_singlepass.json").write_text(
        json.dumps(single_result, indent=2))
    (OUT_DIR / f"slide_{slide_num:02d}_description.txt").write_text(description)

    report = format_diff_report(
        slide_num, diff, multi_cost, single_cost, description)
    (OUT_DIR / f"slide_{slide_num:02d}_diff.txt").write_text(report)

    print(report)

    return {
        "slide":        slide_num,
        "multi_cost":   multi_cost,
        "single_cost":  single_cost,
        "total_keys":   diff["total_keys"],
        "matches":      diff["matches"],
        "match_pct":    diff["match_pct"],
        "mismatches":   diff["mismatches"],
        "only_multi":   diff["only_multi"],
        "only_single":  diff["only_single"],
    }


def main():
    ap = argparse.ArgumentParser(
        description="Compare single-pass vs multi-pass extraction accuracy"
    )
    ap.add_argument("pdf",            help="Path to CFO presentation PDF")
    ap.add_argument("--slide",  type=int, help="Single slide number to test")
    ap.add_argument("--slides", default=None,
                    help="Comma-separated slide numbers, e.g. 3,6,8")
    ap.add_argument("--all-visual", action="store_true",
                    help="Test all slides (skips title/agenda slides)")
    ap.add_argument("--out-dir", default="test_results",
                    help="Output directory (default: test_results/)")
    args = ap.parse_args()

    global OUT_DIR
    OUT_DIR = Path(args.out_dir)

    if not Path(args.pdf).exists():
        sys.exit(f"PDF not found: {args.pdf}")

    # Determine slides to test
    pdf     = pdfium.PdfDocument(args.pdf)
    n_pages = len(pdf)

    if args.slide:
        slides = [args.slide]
    elif args.slides:
        slides = [int(s.strip()) for s in args.slides.split(",")]
    elif args.all_visual:
        slides = list(range(1, n_pages + 1))
    else:
        ap.print_help()
        sys.exit("\nSpecify --slide N, --slides 3,6,8, or --all-visual")

    client  = genai.Client()
    results = []

    for slide_num in slides:
        if slide_num < 1 or slide_num > n_pages:
            print(f"  Slide {slide_num} out of range (PDF has {n_pages} pages), skipping")
            continue
        try:
            r = test_slide(client, args.pdf, slide_num)
            results.append(r)
        except Exception as e:
            print(f"  ❌ slide {slide_num} failed: {e}")
            results.append({"slide": slide_num, "error": str(e)})

    # Summary
    if len(results) > 1:
        print(f"\n{'='*70}")
        print("SUMMARY ACROSS ALL TESTED SLIDES")
        print(f"{'='*70}")
        print(f"{'Slide':<8} {'Keys':<8} {'Match%':<10} {'Mismatches':<14} {'Cost Multi':<14} {'Cost Single'}")
        print(f"{'─'*70}")
        for r in results:
            if "error" in r:
                print(f"{r['slide']:<8} ERROR: {r['error'][:40]}")
            else:
                winner = "← single better" if r["mismatches"] > 0 and r["only_single"] == 0 else ""
                print(f"{r['slide']:<8} {r['total_keys']:<8} {r['match_pct']:<10} "
                      f"{r['mismatches']:<14} ${r['multi_cost']:.5f}      "
                      f"${r['single_cost']:.5f}  {winner}")

        good = [r for r in results if "error" not in r]
        if good:
            avg_match = sum(r["match_pct"] for r in good) / len(good)
            total_multi_cost  = sum(r["multi_cost"] for r in good)
            total_single_cost = sum(r["single_cost"] for r in good)
            print(f"\nAverage match%:     {avg_match:.1f}%")
            print(f"Total multi cost:   ${total_multi_cost:.5f}")
            print(f"Total single cost:  ${total_single_cost:.5f}")
            cost_diff = total_multi_cost - total_single_cost
            print(f"Cost difference:    ${abs(cost_diff):.5f} "
                  f"({'multi more expensive' if cost_diff > 0 else 'single more expensive'})")

        OUT_DIR.mkdir(exist_ok=True)
        (OUT_DIR / "summary.json").write_text(json.dumps(results, indent=2))
        print(f"\nResults saved to {OUT_DIR}/")


if __name__ == "__main__":
    main()
