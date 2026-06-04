# FinDocIQ — Unified Financial Document Extraction Pipeline
## Spec for Claude Code Implementation

---

## Overview

A three-pass pipeline that extracts structured data from bank CFO
presentation slides (DBS, OCBC, UOB) into formatted Excel workbooks.

Works for any visual element type — waterfall, stacked bar, text table,
trend line, KPI grid, pie/donut — without hardcoded layout logic.

**Target cost:** <$0.01 per slide (Gemini 2.5 Flash, thinking_budget=0)

---

## Architecture

```
PDF page
  ↓
render_page() → high-res PNG (3× scale, pypdfium2)
  ↓
PASS 1: describe_slide(image) → description.txt
  ↓
PASS 2: extract_slide(image, description) → datapoints.json
  ↓
validate_extraction(datapoints) → errors[]
  ↓ (if errors: correction_pass → retry once)
PASS 3: render_to_excel(datapoints, ws) → formatted sheet
  ↓
save workbook + update Contents sheet
```

---

## File Structure

```
financial_extractor/
├── extract.py          # main entry point
├── passes/
│   ├── pass1_describe.py
│   ├── pass2_extract.py
│   └── pass3_render.py
├── validation.py
├── models.py           # Pydantic DataPoint schema
├── bank_config.py      # per-bank colours, names
├── utils.py            # render_page, coerce, tab_name, cost tracking
└── outputs/
    ├── audit/
    │   └── {bank}_{doc}/
    │       └── slide_{N}/
    │           ├── description.txt
    │           ├── datapoints.json
    │           └── meta.json          # tokens, cost, validation result
    └── {bank}_{doc}.xlsx
```

---

## Models (models.py)

```python
from pydantic import BaseModel, field_validator
from typing import Any

class DataPoint(BaseModel):
    # Identity
    slide:        int
    element_idx:  int            # 0-based index within slide
    element_type: str            # "waterfall" | "text_table" | "stacked_bar"
                                 # "trend_line" | "kpi_grid" | "pie" | "other"
    element_title: str           # title printed above the element

    # Value
    series:       str            # row label / component / segment name
    period:       str | None     # "1Q26", "Mar-26", "FY25"; null for KPI grids
    value:        str            # verbatim as printed: "5,948" "(5)" "-244" "1.82%"
    value_num:    float | None   # parsed numeric; null if non-numeric
    unit:         str            # "S$m" "S$b" "%" "bps" "x" ""

    # Semantic metadata — Gemini fills from Pass 1 understanding
    row_type:     str            # "data" | "total" | "sub" | "start" | "end" | "bridge" | "note"
    level:        int            # 0=total/header 1=primary 2=sub-item
    parent:       str | None     # parent series label for level>1 items
    group:        str | None     # grouping label e.g. "Commercial book"
    sign:         str | None     # "+" | "-" | null (explicit for bridge components)
    order:        int            # left-to-right / top-to-bottom position on slide

    # Additional columns — Gemini adds freely
    # e.g. pct_change, pct_change_label, qoq_pct, yoy_pct
    # stored as extra_fields dict so schema stays flexible
    extra_fields: dict[str, Any] = {}

    # Provenance
    source:       str            # "table" (high confidence) | "chart" (verify)
    bank:         str            # "DBS" | "OCBC" | "UOB"
    doc_title:    str
    doc_date:     str
    slide_title:  str

    @field_validator("value_num", mode="before")
    @classmethod
    def parse_numeric(cls, v, info):
        if v is not None:
            return v
        raw = info.data.get("value", "")
        if not raw:
            return None
        s = str(raw).strip()
        neg = s.startswith("(") and s.endswith(")")
        try:
            num = float(s[1:-1].replace(",","") if neg else
                        s.replace(",","").replace("+","").rstrip("%"))
            return -num if neg else num
        except ValueError:
            return None


class SlideResult(BaseModel):
    slide:       int
    slide_title: str
    data_points: list[DataPoint]
    validation:  list[str]       # error messages; empty = passed
    cost_usd:    float
```

---

## Pass 1 — describe_slide (passes/pass1_describe.py)

**Goal:** Understand the slide before touching values. Commit to what
each element is and how to read it. Save as plain text.

**Model:** gemini-2.5-flash, thinking_budget=0, NO response schema
**Input:** image only
**Output:** plain text description, saved to audit/slide_N/description.txt

```python
PASS1_PROMPT = """
Examine this bank CFO presentation slide carefully.

List every distinct data element on the slide (tables, charts, KPI boxes).
For each element, describe:

1. TYPE: what kind of visual is it?
   (text_table | waterfall | stacked_bar | trend_line | kpi_grid | pie | other)

2. TITLE: the label printed above it (verbatim)

3. STRUCTURE:
   - text_table: how many rows, how many columns, what are the column headers
   - waterfall: how many bars total, what is the opening bar label,
     what is the closing bar label, what does the colour legend say
     (e.g. "green = positive, red = negative"), are percentage labels
     printed on bars and what do they represent (YoY? QoQ?)
   - stacked_bar: how many time periods, how many stack components,
     what are the period labels, what are the component labels
   - trend_line: how many series, how many periods, what are the series names
   - kpi_grid: how many KPIs, what are the labels
   - pie/donut: how many segments, what do they represent

4. UNITS: what unit are values in (S$m, S$b, %, bps, etc.)

5. VISUAL CONVENTIONS on this specific slide:
   - bold rows = totals?
   - indented rows = sub-items?
   - shaded/grey cells = not applicable?
   - bracket groupings?
   - any footnotes that change interpretation?

Be specific and complete. This description will guide the extraction step.
Do not extract values yet.
"""
```

---

## Pass 2 — extract_slide (passes/pass2_extract.py)

**Goal:** Extract all values as flat DataPoint records.
Gemini decides the fields based on Pass 1 understanding.
No rigid output schema — use a flexible envelope.

**Model:** gemini-2.5-flash, thinking_budget=0
**Input:** image + Pass 1 description
**Output:** JSON list of DataPoint-compatible dicts

```python
def build_pass2_prompt(description: str, bank: str,
                       slide_num: int, slide_title: str) -> str:
    return f"""
You are extracting financial data from a {bank} CFO presentation slide.

You already described this slide as:
<slide_description>
{description}
</slide_description>

Now extract ALL values from every element you identified.

OUTPUT FORMAT:
Return a JSON object with this structure:
{{
  "slide_title": "{slide_title}",
  "elements": [
    {{
      "element_idx": 0,
      "element_type": "...",      // from your description
      "element_title": "...",     // verbatim title
      "source": "table" or "chart",
      "units": "S$m",
      "data_points": [
        {{
          "series": "...",        // row/bar/segment label verbatim
          "period": "...",        // time period; null if none
          "value": "...",         // VERBATIM as printed - keep commas, %, bps, ()
          "row_type": "...",      // data|total|sub|start|end|bridge|note
          "level": 0,             // 0=total 1=primary 2=sub-item
          "parent": null,         // parent series label if level>1
          "group": null,          // grouping bracket label if shown
          "sign": null,           // "+" or "-" for waterfall bridges only
          "order": 0,             // position on slide top-to-bottom/left-to-right
          // ADD any extra columns that make sense for this element:
          // e.g. "qoq_pct": "(1)", "qoq_pct_label": "QoQ % change"
          //      "yoy_pct": "4",  "yoy_pct_label": "YoY % change"
          //      "pct_change": "+16%", "pct_change_label": "YoY % of fee income"
        }}
      ],
      // For waterfall only — include self-validation:
      "waterfall_check": "2897 - 244 + 207 + ... = 2930"
    }}
  ]
}}

RULES:
- value field: ALWAYS verbatim as printed. Never convert.
  "5,948" not 5948. "(5)" not -5. "1.82%" not 0.0182.
- Waterfall bridges: sign field mandatory. Every bar gets "+" or "-".
  Use the colour legend from your description to determine sign.
- Bold rows in text tables → row_type="total", level=0
- Indented rows → level=2, parent=nearest level-1 label above
- Grey/shaded cells → value="" (empty string)
- "record" badges on slides → add as note row, do not lose this information
- If a value is illegible → value="?"
- Extra columns: add them freely if they exist on the slide.
  Always include a matching _label field explaining what the column means.
"""
```

**Parsing the response:**

```python
def parse_pass2_response(raw: str, bank: str, doc_title: str,
                         doc_date: str, slide_num: int,
                         slide_title: str) -> list[DataPoint]:
    # Strip markdown fences
    clean = raw.strip()
    if clean.startswith("```"):
        clean = clean.split("```")[1].lstrip("json").strip()
        clean = clean.rsplit("```")[0].strip()

    data = json.loads(clean)
    points = []

    for elem in data["elements"]:
        for i, dp in enumerate(elem["data_points"]):
            # Separate known fields from extra fields
            known = {"series", "period", "value", "row_type", "level",
                     "parent", "group", "sign", "order"}
            extra = {k: v for k, v in dp.items() if k not in known}

            points.append(DataPoint(
                slide=slide_num,
                slide_title=slide_title,
                element_idx=elem["element_idx"],
                element_type=elem["element_type"],
                element_title=elem["element_title"],
                source=elem["source"],
                unit=elem.get("units", ""),
                bank=bank,
                doc_title=doc_title,
                doc_date=doc_date,
                order=dp.get("order", i),
                extra_fields=extra,
                **{k: dp.get(k) for k in known}
            ))

    return points
```

---

## Validation (validation.py)

```python
def validate(points: list[DataPoint]) -> list[str]:
    errors = []
    by_element = group_by(points, "element_idx")

    for elem_idx, elem_points in by_element.items():
        elem_type = elem_points[0].element_type
        title = elem_points[0].element_title

        # 1. Waterfall balance check
        if elem_type == "waterfall":
            starts = [p for p in elem_points if p.row_type == "start"]
            ends   = [p for p in elem_points if p.row_type == "end"]
            bridge = [p for p in elem_points if p.row_type == "bridge"]

            if starts and ends and bridge:
                opening = starts[0].value_num or 0
                closing = ends[0].value_num or 0
                total   = sum(p.value_num or 0 for p in bridge)
                delta   = abs(opening + total - closing)
                if delta > 5:
                    errors.append(
                        f"Waterfall '{title}': bridge imbalance "
                        f"{opening:.0f} + {total:.0f} ≠ {closing:.0f} "
                        f"(off by {delta:.0f})"
                    )

            # Check no unsigned bridge values
            unsigned = [p.series for p in bridge
                        if p.sign is None and p.value_num is not None
                        and p.value_num != 0]
            if unsigned:
                errors.append(
                    f"Waterfall '{title}': missing sign on: {unsigned}"
                )

        # 2. No blank series labels
        blank_labels = [p for p in elem_points if not p.series.strip()]
        if blank_labels:
            errors.append(
                f"Element '{title}': {len(blank_labels)} rows with blank labels"
            )

        # 3. No illegible values (unless it's a chart)
        if elem_type == "text_table":
            illegible = [p.series for p in elem_points if p.value == "?"]
            if illegible:
                errors.append(
                    f"Table '{title}': illegible values in rows: {illegible}"
                )

    return errors


def build_correction_prompt(errors: list[str],
                            description: str) -> str:
    return f"""
Your previous extraction had these validation errors:
{chr(10).join(f"  - {e}" for e in errors)}

Using the slide description as reference:
<slide_description>
{description}
</slide_description>

Re-extract ONLY the elements with errors, correcting these specific issues.
Return the same JSON format, only including the affected elements.
"""
```

---

## Pass 3 — render_to_excel (passes/pass3_render.py)

**Goal:** Dumb render. No layout decisions. Just write what Gemini gave us.

```python
# Formatting constants
BRAND_COLOURS = {
    "DBS":  "CC0000",
    "OCBC": "CC0000",
    "UOB":  "1B6EC2",
}
NAVY     = "1F3864"
DARK_GREY= "404040"
MID_GREY = "595959"
WHITE    = "FFFFFF"
YELLOW   = "FFFF00"
TOTAL_BG = "F2F2F2"
NUM_FMT  = '#,##0;(#,##0);"-"'
PCT_FMT  = '0.0%;(0.0%);"-"'


def render_slide(ws, points: list[DataPoint],
                 slide_num: int, slide_title: str,
                 bank: str, doc_title: str, doc_date: str):

    brand = BRAND_COLOURS.get(bank, "404040")
    by_element = group_by_ordered(points, "element_idx")
    max_cols = compute_max_cols(points)

    # ── Row 1: brand banner ──
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1, end_column=max_cols)
    c = ws.cell(1, 1, f"{bank} Group  |  Slide {slide_num}: {slide_title}")
    c.fill = PatternFill("solid", fgColor=brand)
    c.font = Font(bold=True, color=WHITE, size=12)
    ws.row_dimensions[1].height = 24

    # ── Row 2: source line ──
    ws.merge_cells(start_row=2, start_column=1,
                   end_row=2, end_column=max_cols)
    c = ws.cell(2, 1,
        f"Source: {doc_title}"
        + (f", {doc_date}" if doc_date else "")
        + "  |  Units: S$m unless noted")
    c.font = Font(italic=True, color=MID_GREY, size=9)

    cursor = 4

    for elem_idx, elem_points in by_element.items():
        cursor = render_element(ws, elem_points, cursor, brand)
        cursor += 1  # spacer between elements

    # ── Chart source warning ──
    if any(p.source == "chart" for p in points):
        ws.merge_cells(start_row=cursor, start_column=1,
                       end_row=cursor, end_column=max_cols)
        c = ws.cell(cursor, 1,
            "⚠ Yellow cells = values read from chart visual "
            "— verify against source slide before use.")
        c.font = Font(italic=True, color="FF6600", size=8)


def render_element(ws, points: list[DataPoint],
                   cursor: int, brand: str) -> int:

    if not points:
        return cursor

    elem_type  = points[0].element_type
    elem_title = points[0].element_title
    is_chart   = points[0].source == "chart"

    # ── Sub-section title ──
    ws.cell(cursor, 1, elem_title).font = Font(
        bold=True, color=brand, size=10)
    cursor += 1

    # ── Discover columns dynamically from DataPoint fields + extra_fields ──
    # Fixed columns always present
    fixed_cols = ["series"]

    # Discover additional columns from the data itself
    # e.g. if points have "period", add it
    # if points have extra_fields keys like "qoq_pct", add them
    extra_col_names = []
    for p in points:
        for k in p.extra_fields:
            if not k.endswith("_label") and k not in extra_col_names:
                extra_col_names.append(k)

    has_period = any(p.period for p in points)
    has_value  = any(p.value for p in points)

    # Build column list
    all_cols = []
    all_cols.append(("series", get_series_header(points)))
    if has_period:
        all_cols.append(("period", "Period"))
    if has_value:
        all_cols.append(("value", f"Value ({points[0].unit or 'S$m'})"))
    for k in extra_col_names:
        # Use the _label field as the column header if available
        label_key = k + "_label"
        header = next(
            (p.extra_fields[label_key] for p in points
             if label_key in p.extra_fields),
            k  # fallback to key name
        )
        all_cols.append((k, header))

    # ── Column headers ──
    for ci, (_, header) in enumerate(all_cols, 1):
        c = ws.cell(cursor, ci, header)
        c.fill = PatternFill("solid", fgColor=DARK_GREY
                             if ci == 1 else "1F3864")
        c.font = Font(bold=True, color=WHITE, size=10)
        c.alignment = Alignment(horizontal="center",
                                vertical="center", wrap_text=True)
    ws.row_dimensions[cursor].height = 20
    cursor += 1

    # ── Data rows ──
    for p in sorted(points, key=lambda x: x.order):
        is_total = p.row_type in ("total", "start", "end")
        is_sub   = p.level >= 2
        is_bridge= p.row_type == "bridge"

        for ci, (field, _) in enumerate(all_cols, 1):
            if field == "series":
                indent = "    " * max(0, p.level - 1)
                val = indent + p.series
            elif field == "period":
                val = p.period or ""
            elif field == "value":
                val = coerce(p.value)
            else:
                # extra field
                val = coerce(p.extra_fields.get(field, ""))

            cell = ws.cell(cursor, ci, val)

            # ── Formatting ──
            if is_total:
                cell.fill = PatternFill("solid", fgColor=TOTAL_BG)
                cell.font = Font(bold=True, size=10)
            elif is_sub:
                cell.font = Font(italic=True, color=MID_GREY, size=9)
            else:
                cell.font = Font(size=10)

            # Chart yellow highlight
            if (is_chart and field == "value"
                    and cell.value not in (None, "", "-", "?")):
                cell.fill = PatternFill("solid", fgColor=YELLOW)

            # Waterfall sign colour
            if is_bridge and field == "value":
                if isinstance(cell.value, (int, float)):
                    colour = "00B050" if cell.value > 0 else "C00000"
                    cell.font = Font(color=colour, size=10)

            # Number formatting
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right")

        cursor += 1

    # ── Waterfall balance check row ──
    if elem_type == "waterfall":
        starts  = [p for p in points if p.row_type == "start"]
        ends    = [p for p in points if p.row_type == "end"]
        bridges = [p for p in points if p.row_type == "bridge"]
        if starts and ends and bridges:
            opening = starts[0].value_num or 0
            closing = ends[0].value_num or 0
            total   = sum(p.value_num or 0 for p in bridges)
            delta   = abs(opening + total - closing)
            msg = (f"✓ Bridge balances: "
                   f"{opening:,.0f} + {total:+,.0f} = {closing:,.0f}"
                   if delta <= 5 else
                   f"⚠ Bridge off by {delta:.0f} — verify source")
            c = ws.cell(cursor, 1, msg)
            c.font = Font(
                italic=True, size=8,
                color="00B050" if delta <= 5 else "FF0000")
            cursor += 1

    return cursor


def get_series_header(points: list[DataPoint]) -> str:
    """Use unit as label column header if available."""
    unit = points[0].unit if points else ""
    return f"({unit})" if unit else "Label"


def compute_max_cols(points: list[DataPoint]) -> int:
    base = 3  # series + period + value
    extra = max(
        (len(p.extra_fields) // 2  # exclude _label fields
         for p in points),
        default=0
    )
    return base + extra
```

---

## Main Entry Point (extract.py)

```python
"""
Usage:
  python extract.py DBS_1Q26_CFO.pdf
  python extract.py UOB_1Q26_CFO.pdf --bank UOB
  python extract.py OCBC_press_release.pdf --start-slide 3
  python extract.py DBS_1Q26_CFO.pdf --slide 5 --force
"""
import os, sys, json, argparse, time
import pypdfium2 as pdfium
import openpyxl
from google import genai
from google.genai import types

from models import DataPoint, SlideResult
from passes.pass1_describe import describe_slide
from passes.pass2_extract import extract_slide
from passes.pass3_render import render_slide, build_contents
from validation import validate, build_correction_prompt
from bank_config import detect_bank, BANKS
from utils import render_page, tab_name, log_usage

MODEL = "gemini-2.5-flash"
MAX_RETRIES = 3
RETRY_ONCE_ON_VALIDATION_FAIL = True


def process_slide(client, pdf_path, page_num,
                  bank, doc_title, doc_date,
                  audit_dir, force=False):

    label      = f"slide_{page_num:02d}"
    audit_path = os.path.join(audit_dir, label)
    desc_path  = os.path.join(audit_path, "description.txt")
    dp_path    = os.path.join(audit_path, "datapoints.json")
    meta_path  = os.path.join(audit_path, "meta.json")
    os.makedirs(audit_path, exist_ok=True)

    # ── Resume from audit if available ──
    if not force and os.path.exists(dp_path):
        with open(dp_path) as f:
            raw = json.load(f)
        points = [DataPoint(**p) for p in raw]
        print(f"  slide {page_num:02d}  resumed ({len(points)} points)")
        return points, 0.0

    # ── Render page ──
    img_bytes = render_page(pdf_path, page_num, scale=3.0)

    total_cost = 0.0

    # ── Pass 1: Describe ──
    description, cost1 = describe_slide(client, img_bytes, MODEL)
    total_cost += cost1

    with open(desc_path, "w") as f:
        f.write(description)

    if not description.strip():
        print(f"  slide {page_num:02d}  — no content described, skipping")
        return [], total_cost

    # ── Pass 2: Extract ──
    points, cost2 = extract_slide(
        client, img_bytes, description,
        bank, doc_title, doc_date,
        page_num, slide_title="",  # filled in by pass2 from response
        model=MODEL
    )
    total_cost += cost2

    # ── Validate ──
    errors = validate(points)

    # ── Correction pass if validation failed ──
    if errors and RETRY_ONCE_ON_VALIDATION_FAIL:
        print(f"  slide {page_num:02d}  ⚠ validation errors: {errors}")
        correction_prompt = build_correction_prompt(errors, description)
        points_corrected, cost3 = extract_slide(
            client, img_bytes, description,
            bank, doc_title, doc_date,
            page_num, slide_title="",
            model=MODEL,
            correction=correction_prompt
        )
        total_cost += cost3
        errors_after = validate(points_corrected)

        if len(errors_after) < len(errors):
            points = points_corrected
            errors = errors_after
            print(f"  slide {page_num:02d}  ✓ correction improved output")
        else:
            print(f"  slide {page_num:02d}  ⚠ correction did not help, using original")

    # ── Save audit ──
    with open(dp_path, "w") as f:
        json.dump([p.model_dump() for p in points], f, indent=2)
    with open(meta_path, "w") as f:
        json.dump({
            "slide": page_num, "cost_usd": total_cost,
            "n_points": len(points), "validation_errors": errors
        }, f, indent=2)

    status = "✓" if not errors else f"⚠ {len(errors)} error(s)"
    chart_flag = " 📊" if any(p.source == "chart" for p in points) else ""
    print(f"  slide {page_num:02d}  {len(points)} points  "
          f"${total_cost:.4f}  {status}{chart_flag}")

    return points, total_cost


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--bank",        choices=["DBS", "OCBC", "UOB"])
    ap.add_argument("--out",         default=None)
    ap.add_argument("--slide",       type=int)
    ap.add_argument("--start-slide", type=int)
    ap.add_argument("--force",       action="store_true")
    ap.add_argument("--dry-run",     action="store_true")
    ap.add_argument("--doc-date",    default="")
    args = ap.parse_args()

    detected, det_date = detect_bank(args.pdf)
    bank     = args.bank or detected or "DBS"
    doc_date = args.doc_date or det_date or ""
    doc_title= os.path.basename(args.pdf).replace(".pdf", "")
    bank_slug= bank.lower()

    out_path  = args.out or f"outputs/{bank_slug}_extracted.xlsx"
    audit_dir = f"outputs/audit/{bank_slug}_{doc_title}"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    pdf     = pdfium.PdfDocument(args.pdf)
    n_pages = len(pdf)
    pages   = (
        [args.slide] if args.slide else
        [p for p in range(1, n_pages+1)
         if not args.start_slide or p >= args.start_slide]
    )

    print(f"📄 {args.pdf}  ({n_pages} slides)  bank={bank}")
    print(f"   Output → {out_path}")

    if args.dry_run:
        print(f"\nDRY RUN — {len(pages)} slides would be processed")
        return

    client = genai.Client()

    if os.path.exists(out_path):
        wb = openpyxl.load_workbook(out_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    used_names = set(wb.sheetnames)
    index      = []
    total_cost = 0.0

    for pg in pages:
        points, cost = process_slide(
            client, args.pdf, pg,
            bank, doc_title, doc_date,
            audit_dir, force=args.force
        )
        total_cost += cost

        if not points:
            continue

        slide_title = points[0].slide_title if points else f"Slide {pg}"
        sname = tab_name(used_names, pg, slide_title)

        if sname in wb.sheetnames:
            del wb[sname]
        ws = wb.create_sheet(title=sname)

        render_slide(ws, points, pg, slide_title,
                     bank, doc_title, doc_date)

        index.append({
            "slide_num":   pg,
            "slide_title": slide_title,
            "sheet":       sname,
            "n_points":    len(points),
            "has_chart":   any(p.source == "chart" for p in points),
            "has_errors":  bool(validate(points))
        })

        build_contents(wb, index, bank, doc_title, doc_date)
        wb.save(out_path)
        print(f"    💾 saved  |  running cost ≈ ${total_cost:.4f}")

    print(f"\n✅ Done → {out_path}")
    print(f"   Total cost ≈ ${total_cost:.4f}")


if __name__ == "__main__":
    main()
```

---

## Bank Config (bank_config.py)

```python
import re, io
import pypdfium2 as pdfium

BANKS = {
    "DBS":  {
        "institution":  "DBS Group Holdings Ltd",
        "brand_colour": "CC0000",
        "match":        r"\bDBS\b",
    },
    "OCBC": {
        "institution":  "Oversea-Chinese Banking Corporation Limited",
        "brand_colour": "CC0000",
        "match":        r"OCBC|Oversea[- ]?Chinese",
    },
    "UOB":  {
        "institution":  "United Overseas Bank Limited",
        "brand_colour": "1B6EC2",
        "match":        r"\bUOB\b|United Overseas",
    },
}

def detect_bank(pdf_path: str) -> tuple[str | None, str | None]:
    try:
        doc  = pdfium.PdfDocument(pdf_path)
        txt  = doc[0].get_textpage().get_text_range()
        if len(doc) > 1:
            txt += " " + doc[1].get_textpage().get_text_range()
    except Exception:
        txt = ""
    for key, info in BANKS.items():
        if re.search(info["match"], txt, re.I):
            bank = key
            break
    else:
        bank = None
    m = re.search(r"\b(\d{1,2}\s+[A-Za-z]+\s+\d{4})\b", txt)
    return bank, (m.group(1) if m else None)
```

---

## Key Design Decisions (for Claude Code)

1. **Pass 1 has no Pydantic schema** — plain text output only.
   Gemini must not feel schema pressure during visual comprehension.

2. **Pass 2 output is a flexible JSON envelope** — `extra_fields` on
   DataPoint absorbs any columns Gemini adds (qoq_pct, yoy_pct, etc.)
   without breaking the schema.

3. **Pass 3 discovers columns dynamically** — it reads whatever fields
   are present in the DataPoints for that element. No hardcoded layout
   per chart type. Gemini's understanding of the slide determines the
   columns; the renderer just writes them.

4. **Waterfall balance check is arithmetic** — derived from value_num
   fields, not from Gemini's waterfall_check string. But Gemini's
   check is saved in audit for reference.

5. **Audit trail is three files per slide:**
   - `description.txt` — what Gemini saw (Pass 1)
   - `datapoints.json` — what was extracted (Pass 2)
   - `meta.json` — cost, token counts, validation errors

6. **Resume is at DataPoint level** — if `datapoints.json` exists and
   `--force` is not set, skip both Pass 1 and Pass 2. This means a
   partial run (e.g. API timeout on slide 12) resumes from slide 12
   without re-processing slides 1-11.

7. **Correction pass fires once** — if validation fails, one targeted
   correction prompt referencing specific errors. If it doesn't improve,
   proceed with original output and flag in meta.json.

8. **Cost target: <$0.01 per data slide** at Gemini 2.5 Flash rates.
   thinking_budget=0 everywhere. Monitor via meta.json cost field.

---

## Dependencies

```
pypdfium2
openpyxl
pydantic>=2.0
google-genai
```

Install:
```bash
pip install pypdfium2 openpyxl pydantic google-genai
```
