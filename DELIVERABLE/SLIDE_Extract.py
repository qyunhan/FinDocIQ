"""
SLIDE_Extract.py — CFO presentation slide deck → Excel workbook.
Three-pass architecture per PIPELINE_SPEC.md:
  Pass 1: describe_slide  → description.txt  (plain text, no schema)
  Pass 2: extract_slide   → datapoints.json  (flexible DataPoint JSON)
  Validate + correct once if waterfall/label errors found
  Pass 3: render_to_excel → one worksheet per slide

Usage:
  export GEMINI_API_KEY=...
  python3 SLIDE_Extract.py DBS4Q25_CFO_presentation.pdf
  python3 SLIDE_Extract.py DBS4Q25_CFO_presentation.pdf --slide 5
  python3 SLIDE_Extract.py DBS4Q25_CFO_presentation.pdf --start-slide 3
  python3 SLIDE_Extract.py DBS4Q25_CFO_presentation.pdf --dry-run
  python3 SLIDE_Extract.py DBS4Q25_CFO_presentation.pdf --force
"""
from __future__ import annotations
import os, sys, json, io, re, argparse, time
from typing import Any
import pypdfium2 as pdfium
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from pydantic import BaseModel, field_validator
from google import genai
from google.genai import types

# ---------------------------------------------------------------------------
MODEL      = "gemini-2.5-flash"
IMAGE_SCALE = 3.0

BANKS = {
    "DBS":  {"institution": "DBS Group Holdings Ltd",
             "brand": "CC0000", "match": r"\bDBS\b"},
    "OCBC": {"institution": "Oversea-Chinese Banking Corporation Limited",
             "brand": "CC0000", "match": r"OCBC|Oversea[- ]?Chinese"},
    "UOB":  {"institution": "United Overseas Bank Limited",
             "brand": "1B6EC2", "match": r"\bUOB\b|United Overseas"},
}

# Gemini 2.5 Flash pricing
INPUT_PRICE_PER_M  = 0.30
OUTPUT_PRICE_PER_M = 2.50

NAVY      = "1F3864"
DARK_GREY = "404040"
MID_GREY  = "595959"
WHITE     = "FFFFFF"
YELLOW    = "FFFF00"
TOTAL_BG  = "F2F2F2"
NUM_FMT   = '#,##0;(#,##0);"-"'

# Gemini sometimes uses synonyms; normalise before DataPoint construction.
ROW_TYPE_ALIASES: dict[str, str] = {
    "header":         "total",
    "section_header": "total",
    "sub_header":     "sub",
    "subtotal":       "total",
    "opening":        "start",
    "closing":        "end",
    "component":      "bridge",
    "commentary":     "note",
    "footnote":       "note",
}

_run_usage: dict = {"calls": 0, "prompt": 0, "output": 0, "cost": 0.0}

# ===========================================================================
# DATA MODEL
# ===========================================================================
class DataPoint(BaseModel):
    # Identity
    slide:         int
    element_idx:   int
    element_type:  str = "other"
    element_title: str = ""

    # Value
    series:    str = ""
    period:    str | None = None
    value:     str = ""
    value_num: float | None = None
    unit:      str = ""

    # Semantic metadata
    row_type: str = "data"
    level:    int = 1
    parent:   str | None = None
    group:    str | None = None
    sign:     str | None = None   # "+" or "-" for waterfall bridges
    order:    int = 0

    # Dynamic extra columns (qoq_pct, yoy_pct, etc.)
    extra_fields: dict[str, Any] = {}

    # Provenance
    source:     str   # "table" | "chart"
    bank:       str
    doc_title:  str
    doc_date:   str
    slide_title: str

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
            core = s[1:-1] if neg else s
            num = float(core.replace(",", "").replace("+", "").rstrip("%"))
            return -num if neg else num
        except ValueError:
            return None


# ===========================================================================
# PASS 1 PROMPT
# ===========================================================================
PASS1_PROMPT = """Examine this bank CFO presentation slide carefully.

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
Do not extract values yet."""


# ===========================================================================
# PASS 2 PROMPT
# ===========================================================================
def build_pass2_prompt(description: str, bank: str) -> str:
    return f"""You are extracting financial data from a {bank} CFO presentation slide.

You already described this slide as:
<slide_description>
{description}
</slide_description>

Now extract ALL values from every element you identified.

OUTPUT FORMAT:
Return a JSON object with this structure:
{{
  "slide_title": "...",
  "elements": [
    {{
      "element_idx": 0,
      "element_type": "...",
      "element_title": "...",
      "source": "table",
      "units": "S$m",
      "self_check": "2897 - 244 + 207 = 2860",
      "data_points": [
        {{
          "series": "...",
          "period": null,
          "value": "...",
          "row_type": "data",
          "level": 1,
          "parent": null,
          "group": null,
          "sign": null,
          "order": 0
        }}
      ]
    }}
  ]
}}

RULES:
- value field: ALWAYS verbatim as printed. Never convert.
  "5,948" not 5948. "(5)" not -5. "1.82%" not 0.0182.
- Waterfall bridges: sign field mandatory ("+" or "-") for every bar.
  Use the colour legend from your description to determine sign.
- Bold rows in text tables → row_type="total", level=0
- Indented rows → level=2, parent=nearest level-1 label above
- Grey/shaded cells → value="" (empty string)
- "record" badges → add as a row with row_type="note"
- If a value is illegible → value="?"
- Extra columns: add freely if they exist on the slide
  (e.g. "qoq_pct": "(1)", "qoq_pct_label": "QoQ % change").
  Always include a matching _label field for every extra column.
- slide_title: main heading of the slide verbatim
- self_check: write the arithmetic for any element where values should sum to a
  total (waterfall: "2897 - 244 + 207 = 2930"; stacked bar total: "724 + 649 = 1373";
  pie: "176 + 77 + 49 = 302"). Set null if no summation relationship exists.

Return ONLY the JSON object, no markdown fences."""


# ===========================================================================
# CORRECTION PROMPT
# ===========================================================================
def build_correction_prompt(errors: list[str], description: str) -> str:
    return f"""Your previous extraction had these validation errors:
{chr(10).join(f"  - {e}" for e in errors)}

Using the slide description as reference:
<slide_description>
{description}
</slide_description>

For any self_check arithmetic failure: re-examine the values or signs in the
affected element so the equation holds. Write the corrected self_check.

Re-extract ONLY the elements with errors, correcting these specific issues.
Return the same JSON format, only including the affected elements.
Return ONLY the JSON object, no markdown fences."""


# ===========================================================================
# HELPERS
# ===========================================================================
def render_page(pdf_path: str, page_1based: int, scale: float = IMAGE_SCALE) -> bytes:
    src = pdfium.PdfDocument(pdf_path)
    pil = src[page_1based - 1].render(scale=scale).to_pil()
    buf = io.BytesIO()
    pil.save(buf, format="PNG")
    return buf.getvalue()


def detect_bank(pdf_path: str) -> tuple[str | None, str | None]:
    try:
        pdf = pdfium.PdfDocument(pdf_path)
        txt = pdf[0].get_textpage().get_text_range()
        if len(pdf) > 1:
            txt += " " + pdf[1].get_textpage().get_text_range()
    except Exception:
        txt = ""
    key = None
    for k, info in BANKS.items():
        if re.search(info["match"], txt, re.I):
            key = k
            break
    m = re.search(r"\b(\d{1,2}\s+[A-Za-z]+\s+\d{4})\b", txt)
    return key, (m.group(1) if m else None)


def call_gemini(client, prompt_parts: list, *, text_only: bool = False) -> tuple[str, dict]:
    """Single Gemini call with retry. Returns (text, usage_rec)."""
    config_kwargs: dict = {"temperature": 0.0}
    if not text_only:
        config_kwargs["response_mime_type"] = "application/json"
    try:
        config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_budget=0)
    except Exception:
        pass
    config = types.GenerateContentConfig(**config_kwargs)

    resp = None
    for attempt in range(3):
        try:
            resp = client.models.generate_content(
                model=MODEL,
                contents=prompt_parts,
                config=config,
            )
            break
        except Exception as e:
            if any(c in str(e) for c in ("503", "429", "UNAVAILABLE", "RESOURCE_EXHAUSTED")):
                wait = 15 * (2 ** attempt)
                print(f"      ⏳ {e.__class__.__name__} — waiting {wait}s")
                time.sleep(wait)
            else:
                raise

    if resp is None:
        raise RuntimeError("Gemini failed after 3 retries")

    text = (resp.text or "").strip()
    um   = getattr(resp, "usage_metadata", None)
    pt   = getattr(um, "prompt_token_count", None) or 0
    ot   = getattr(um, "candidates_token_count", None) or 0
    cost = (pt / 1e6 * INPUT_PRICE_PER_M) + (ot / 1e6 * OUTPUT_PRICE_PER_M)
    usage = {"prompt_tokens": pt, "output_tokens": ot, "est_cost_usd": round(cost, 6)}
    _run_usage["calls"]  += 1
    _run_usage["prompt"] += pt
    _run_usage["output"] += ot
    _run_usage["cost"]   += cost
    return text, usage


def img_part(img_bytes: bytes) -> types.Part:
    return types.Part.from_bytes(
        data=img_bytes, mime_type="image/png",
        media_resolution=types.PartMediaResolutionLevel.MEDIA_RESOLUTION_HIGH,
    )


def strip_fences(raw: str) -> str:
    s = raw.strip()
    if s.startswith("```"):
        s = s.split("```", 1)[1].lstrip("json").strip()
        s = s.rsplit("```", 1)[0].strip()
    return s


def _auto_label(key: str) -> str:
    """Generate a human-readable label from a snake_case key."""
    return key.replace("_", " ").title()


def parse_pass2(raw: str, bank: str, doc_title: str, doc_date: str,
                slide_num: int) -> tuple[list[DataPoint], str, list[str]]:
    """Returns (points, slide_title, self_checks)."""
    data = json.loads(strip_fences(raw))
    slide_title  = data.get("slide_title", "")
    points: list[DataPoint] = []
    self_checks: list[str] = []
    known = {"series", "period", "value", "row_type", "level",
             "parent", "group", "sign", "order"}

    for elem in data.get("elements", []):
        sc = elem.get("self_check")
        if sc:
            self_checks.append(f"[{elem.get('element_title', '')}] {sc}")

        for i, dp in enumerate(elem.get("data_points", [])):
            # Normalise row_type
            rt = str(dp.get("row_type") or "data")
            rt = ROW_TYPE_ALIASES.get(rt, rt)

            extra = {k: v for k, v in dp.items() if k not in known}

            # Auto-generate _label for any extra key missing one
            for k in list(extra):
                if not k.endswith("_label") and (k + "_label") not in extra:
                    extra[k + "_label"] = _auto_label(k)

            known_vals = {k: dp.get(k) for k in known}
            known_vals["row_type"] = rt
            # Coerce nulls to safe defaults so Pydantic doesn't reject them
            if known_vals.get("order") is None:
                known_vals["order"] = i
            if known_vals.get("level") is None:
                known_vals["level"] = 1
            if known_vals.get("series") is None:
                known_vals["series"] = ""
            if known_vals.get("value") is None:
                known_vals["value"] = ""

            points.append(DataPoint(
                slide=slide_num,
                slide_title=slide_title or "",
                element_idx=elem["element_idx"],
                element_type=elem.get("element_type") or "other",
                element_title=elem.get("element_title") or "",
                source=elem.get("source") or "table",
                unit=elem.get("units") or "",
                bank=bank,
                doc_title=doc_title,
                doc_date=doc_date,
                extra_fields=extra,
                **known_vals,
            ))
    return points, slide_title, self_checks


def merge_correction(original: list[DataPoint],
                     corrected: list[DataPoint]) -> list[DataPoint]:
    """Replace elements from original with corrected versions where idx matches."""
    corrected_idxs = {p.element_idx for p in corrected}
    merged = [p for p in original if p.element_idx not in corrected_idxs]
    merged.extend(corrected)
    merged.sort(key=lambda p: (p.element_idx, p.order))
    return merged


# ===========================================================================
# PDFPLUMBER CROSS-CHECK
# ===========================================================================
def _plumber_values(pdf_path: str, page_0based: int) -> set[str]:
    """Extract all numeric strings from native text layer via pdfplumber."""
    if not _HAS_PDFPLUMBER:
        return set()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_0based >= len(pdf.pages):
                return set()
            text = pdf.pages[page_0based].extract_text() or ""
        # collect tokens that look like financial figures
        return {t for t in re.findall(r'[\d,]+(?:\.\d+)?', text) if len(t) > 1}
    except Exception:
        return set()


def _is_clean_table_page(pdf_path: str, page_0based: int) -> bool:
    """Return True if the page has a meaningful native text layer (not a pure chart slide)."""
    if not _HAS_PDFPLUMBER:
        return False
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_0based >= len(pdf.pages):
                return False
            tables = pdf.pages[page_0based].extract_tables()
            return bool(tables)
    except Exception:
        return False


def crosscheck_with_textlayer(points: list[DataPoint],
                              pdf_path: str, page_num: int) -> list[DataPoint]:
    """
    For text_table elements on pages that have a native table layer:
    mark any value NOT found in the text layer as source="unverified".
    unverified renders yellow the same as chart values.
    """
    if not _HAS_PDFPLUMBER:
        return points

    page_0 = page_num - 1
    if not _is_clean_table_page(pdf_path, page_0):
        return points

    native_vals = _plumber_values(pdf_path, page_0)
    if not native_vals:
        return points

    updated = []
    for p in points:
        if p.element_type == "text_table" and p.source == "table":
            # Strip commas/signs from the extracted value for comparison
            core = re.sub(r'[,\+\(\)]', '', p.value).strip().lstrip("-")
            if core and core not in native_vals:
                p = p.model_copy(update={"source": "unverified"})
        updated.append(p)
    return updated


# ===========================================================================
# VALIDATION
# ===========================================================================
def validate_self_check(self_check: str | None) -> str | None:
    if not self_check:
        return None
    try:
        lhs, rhs = self_check.split("=")
        expected = float(rhs.strip().replace(",", ""))
        actual   = eval(lhs.strip().replace(",", ""), {"__builtins__": {}})
        delta    = abs(actual - expected)
        if delta > 5:
            return (f"Self-check failed: {self_check.strip()} "
                    f"(off by {delta:.0f})")
        return None
    except Exception:
        return f"Self-check unparseable: {self_check}"


def validate(points: list[DataPoint], self_checks: list[str] | None = None) -> list[str]:
    errors: list[str] = []

    # Generic arithmetic self-checks written by Gemini during extraction
    for sc in (self_checks or []):
        err = validate_self_check(sc)
        if err:
            errors.append(err)

    # Structural checks (element-type agnostic)
    by_elem: dict[int, list[DataPoint]] = {}
    for p in points:
        by_elem.setdefault(p.element_idx, []).append(p)

    for pts in by_elem.values():
        title = pts[0].element_title
        etype = pts[0].element_type

        blank = [p for p in pts if not p.series.strip()]
        if blank:
            errors.append(f"Element '{title}': {len(blank)} rows with blank labels")

        if etype == "text_table":
            illegible = [p.series for p in pts if p.value == "?"]
            if illegible:
                errors.append(f"Table '{title}': illegible values in rows: {illegible}")

    return errors


# ===========================================================================
# PASS 3 — EXCEL RENDER
# ===========================================================================
def coerce(v: Any) -> Any:
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "-", "–", "—", "n.m.", "nm", "NA", "N/A", "?"):
        return s
    if s.endswith("%") or s.endswith("bps") or s.endswith("x"):
        return s
    t = s.replace(",", "")
    neg = t.startswith("(") and t.endswith(")")
    core = t[1:-1] if neg else t
    try:
        num = float(core.replace("+", ""))
        if neg:
            num = -num
        return int(num) if num == int(num) else num
    except ValueError:
        return s


def tab_name(used: set, slide_num: int, slide_title: str) -> str:
    prefix = f"{slide_num:02d}"
    short  = re.sub(r'[\\/:*?\[\]]', '', slide_title).strip()[:20] or "Slide"
    base   = f"{prefix} - {short}"[:31]
    name, i = base, 2
    while name in used:
        suffix = f" ({i})"
        name = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(name)
    return name


def _hdr(cell, dark: bool = False):
    cell.fill = PatternFill("solid", fgColor=DARK_GREY if dark else NAVY)
    cell.font = Font(bold=True, color=WHITE, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def render_slide(ws, points: list[DataPoint], slide_num: int, slide_title: str,
                 bank: str, doc_title: str, doc_date: str, brand: str):
    by_elem: dict[int, list[DataPoint]] = {}
    for p in points:
        by_elem.setdefault(p.element_idx, []).append(p)

    # Compute max columns across all elements for banner merge
    max_cols = _compute_max_cols(points)

    # Row 1 — brand banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
    c = ws.cell(1, 1, f"{bank} Group  |  Slide {slide_num}: {slide_title}")
    c.fill = PatternFill("solid", fgColor=brand)
    c.font = Font(bold=True, color=WHITE, size=12)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 24

    # Row 2 — source line
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_cols)
    c = ws.cell(2, 1,
        f"Source: {doc_title}"
        + (f", {doc_date}" if doc_date else "")
        + "  |  Units: S$m unless noted")
    c.font = Font(italic=True, color=MID_GREY, size=9)

    cursor = 4

    for elem_idx in sorted(by_elem):
        cursor = _render_element(ws, by_elem[elem_idx], cursor, brand)
        cursor += 1  # spacer

    # Chart / unverified source warning
    if any(p.source in ("chart", "unverified") for p in points):
        ws.merge_cells(start_row=cursor, start_column=1,
                       end_row=cursor, end_column=max_cols)
        c = ws.cell(cursor, 1,
            "⚠ Yellow cells = values read from chart or not confirmed by text layer "
            "— verify against source slide before use.")
        c.font = Font(italic=True, color="FF6600", size=8)

    # Column widths
    ws.column_dimensions[get_column_letter(1)].width = 48
    for i in range(2, max_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14


def _compute_max_cols(points: list[DataPoint]) -> int:
    base  = 3  # series + period + value
    extra = max(
        (len([k for k in p.extra_fields if not k.endswith("_label")])
         for p in points),
        default=0,
    )
    return max(base + extra, 4)


def _render_element(ws, points: list[DataPoint], cursor: int, brand: str) -> int:
    if not points:
        return cursor

    elem_title = points[0].element_title

    # Sub-section title
    ws.cell(cursor, 1, elem_title).font = Font(bold=True, color=brand, size=10)
    cursor += 1

    # Discover columns
    extra_keys: list[str] = []
    for p in points:
        for k in p.extra_fields:
            if not k.endswith("_label") and k not in extra_keys:
                extra_keys.append(k)

    has_period = any(p.period for p in points)
    has_value  = any(p.value for p in points)

    all_cols: list[tuple[str, str]] = []
    unit = points[0].unit or "S$m"
    all_cols.append(("series", f"({unit})"))
    if has_period:
        all_cols.append(("period", "Period"))
    if has_value:
        all_cols.append(("value", f"Value ({unit})"))
    for k in extra_keys:
        label_key = k + "_label"
        header = next(
            (p.extra_fields[label_key] for p in points if label_key in p.extra_fields),
            k,
        )
        all_cols.append((k, header))

    # Column headers
    for ci, (_, header) in enumerate(all_cols, 1):
        _hdr(ws.cell(cursor, ci, header), dark=(ci == 1))
    ws.row_dimensions[cursor].height = 20
    cursor += 1

    # Data rows
    for p in sorted(points, key=lambda x: x.order):
        is_total  = p.row_type in ("total", "start", "end")
        is_sub    = p.level >= 2
        is_bridge = p.row_type == "bridge"

        for ci, (field, _) in enumerate(all_cols, 1):
            if field == "series":
                indent = "    " * max(0, p.level - 1)
                val = indent + p.series
            elif field == "period":
                val = p.period or ""
            elif field == "value":
                val = coerce(p.value)
            else:
                val = coerce(p.extra_fields.get(field, ""))

            cell = ws.cell(cursor, ci, val)

            if is_total:
                cell.fill = PatternFill("solid", fgColor=TOTAL_BG)
                cell.font = Font(bold=True, size=10)
            elif is_sub:
                cell.font = Font(italic=True, color=MID_GREY, size=9)
            else:
                cell.font = Font(size=10)

            if (p.source in ("chart", "unverified") and field == "value"
                    and cell.value not in (None, "", "-", "?")):
                cell.fill = PatternFill("solid", fgColor=YELLOW)

            if is_bridge and field == "value" and isinstance(cell.value, (int, float)):
                colour = "00B050" if cell.value > 0 else "C00000"
                cell.font = Font(color=colour, size=10)

            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right")

        cursor += 1

    return cursor


def build_contents(wb, index: list[dict], bank: str, doc_title: str,
                   doc_date: str, brand: str):
    if "Contents" in wb.sheetnames:
        wb.remove(wb["Contents"])
    ws = wb.create_sheet("Contents", 0)

    ws.merge_cells("A1:E1")
    c = ws.cell(1, 1, BANKS[bank]["institution"] if bank in BANKS else bank)
    c.fill = PatternFill("solid", fgColor=brand)
    c.font = Font(bold=True, color=WHITE, size=14)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:E2")
    c = ws.cell(2, 1, f"{doc_title}{' — ' + doc_date if doc_date else ''}")
    c.font = Font(italic=True, size=11)
    c.alignment = Alignment(horizontal="center")

    for ci, h in enumerate(["Slide", "Title", "Points", "Has Charts", "Sheet"], 1):
        _hdr(ws.cell(4, ci, h))

    for r, e in enumerate(sorted(index, key=lambda x: x["slide_num"]), start=5):
        ws.cell(r, 1, e["slide_num"]).alignment = Alignment(horizontal="center")
        ws.cell(r, 2, e["slide_title"])
        ws.cell(r, 3, e["n_points"]).alignment = Alignment(horizontal="center")
        has_chart = "yes" if e.get("has_chart") else "no"
        c = ws.cell(r, 4, has_chart)
        c.alignment = Alignment(horizontal="center")
        if has_chart == "yes":
            c.font = Font(color="FF6600")
        link = ws.cell(r, 5, e["sheet"])
        link.hyperlink = f"#'{e['sheet']}'!A1"
        link.font = Font(color="0000CC", underline="single")

    for col, w in {1: 8, 2: 60, 3: 10, 4: 12, 5: 32}.items():
        ws.column_dimensions[get_column_letter(col)].width = w


# ===========================================================================
# SLIDE PROCESSOR
# ===========================================================================
def process_slide(client, pdf_path: str, page_num: int,
                  bank: str, doc_title: str, doc_date: str,
                  audit_dir: str,
                  force: bool = False) -> tuple[list[DataPoint], float]:

    audit_path = os.path.join(audit_dir, f"slide_{page_num:02d}")
    desc_path  = os.path.join(audit_path, "description.txt")
    dp_path    = os.path.join(audit_path, "datapoints.json")
    meta_path  = os.path.join(audit_path, "meta.json")
    os.makedirs(audit_path, exist_ok=True)

    # Resume from audit
    if not force and os.path.exists(dp_path):
        raw = json.load(open(dp_path))
        points = [DataPoint(**p) for p in raw]
        print(f"  slide {page_num:02d}  resumed ({len(points)} points)")
        return points, 0.0

    img_bytes = render_page(pdf_path, page_num)
    total_cost = 0.0
    usages: list[dict] = []

    # ── Pass 1: Describe ──
    desc_text, u1 = call_gemini(
        client,
        [img_part(img_bytes), PASS1_PROMPT],
        text_only=True,
    )
    total_cost += u1["est_cost_usd"]
    usages.append({"pass": 1, **u1})

    with open(desc_path, "w") as f:
        f.write(desc_text)

    if not desc_text.strip():
        print(f"  slide {page_num:02d}  — nothing described, skipped")
        return [], total_cost

    # ── Pass 2: Extract ──
    p2_prompt = build_pass2_prompt(desc_text, bank)
    raw2, u2 = call_gemini(client, [img_part(img_bytes), p2_prompt])
    total_cost += u2["est_cost_usd"]
    usages.append({"pass": 2, **u2})

    try:
        points, _, self_checks = parse_pass2(raw2, bank, doc_title, doc_date, page_num)
    except Exception as e:
        print(f"  slide {page_num:02d}  ❌ parse error: {e}")
        return [], total_cost

    # ── pdfplumber cross-check ──
    points = crosscheck_with_textlayer(points, pdf_path, page_num)

    # ── Validate ──
    errors = validate(points, self_checks)

    # ── Correction pass (once) ──
    if errors:
        print(f"  slide {page_num:02d}  ⚠ validation: {errors}")
        corr_prompt = build_correction_prompt(errors, desc_text)
        try:
            raw3, u3 = call_gemini(client, [img_part(img_bytes), corr_prompt])
            total_cost += u3["est_cost_usd"]
            usages.append({"pass": "2c", **u3})
            corrected, _, self_checks_c = parse_pass2(raw3, bank, doc_title, doc_date, page_num)
            corrected = crosscheck_with_textlayer(corrected, pdf_path, page_num)
            errors_after = validate(corrected, self_checks_c)
            if len(errors_after) < len(errors):
                points = merge_correction(points, corrected)
                errors = errors_after
                self_checks = self_checks_c
                print(f"  slide {page_num:02d}  ✓ correction improved output")
            else:
                print(f"  slide {page_num:02d}  ⚠ correction did not help, using original")
        except Exception as e:
            print(f"  slide {page_num:02d}  ⚠ correction failed: {e}")

    # ── Save audit ──
    with open(dp_path, "w") as f:
        json.dump([p.model_dump() for p in points], f, indent=2)
    with open(meta_path, "w") as f:
        json.dump({"slide": page_num, "cost_usd": total_cost,
                   "n_points": len(points), "validation_errors": errors,
                   "self_checks": self_checks, "usages": usages}, f, indent=2)

    status     = "✓" if not errors else f"⚠ {len(errors)} err"
    unverified = sum(1 for p in points if p.source == "unverified")
    flags      = (" 📊" if any(p.source == "chart" for p in points) else "")
    flags     += (f" ⚡{unverified}unverified" if unverified else "")
    in_t  = sum(u.get("prompt_tokens", 0) for u in usages)
    out_t = sum(u.get("output_tokens", 0) for u in usages)
    print(f"  slide {page_num:02d}  {len(points)} pts  "
          f"[{in_t}in/{out_t}out]  ${total_cost:.4f}  {status}{flags}")

    return points, total_cost


# ===========================================================================
# MAIN
# ===========================================================================
def main():
    ap = argparse.ArgumentParser(description="CFO slide deck PDF → Excel (3-pass)")
    ap.add_argument("pdf")
    ap.add_argument("--out",         default=None)
    ap.add_argument("--slide",       type=int, help="extract this slide only (1-based)")
    ap.add_argument("--start-slide", type=int, help="resume from this slide number")
    ap.add_argument("--bank",        choices=list(BANKS))
    ap.add_argument("--doc-date",    default="")
    ap.add_argument("--force",       action="store_true",
                    help="re-extract even if audit exists")
    ap.add_argument("--dry-run",     action="store_true",
                    help="list slides without calling API")
    args = ap.parse_args()

    if not os.path.exists(args.pdf):
        sys.exit(f"PDF not found: {args.pdf}")

    detected, det_date = detect_bank(args.pdf)
    bank      = args.bank or detected or "DBS"
    doc_date  = args.doc_date or det_date or ""
    doc_title = os.path.basename(args.pdf).replace(".pdf", "")
    bank_slug = bank.lower()
    brand     = BANKS[bank]["brand"]

    out_path  = args.out or f"outputs/{bank_slug}_slides.xlsx"
    audit_dir = f"outputs/audit/slides/{bank_slug}_{doc_title}"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    pdf     = pdfium.PdfDocument(args.pdf)
    n_pages = len(pdf)
    if args.slide:
        pages = [args.slide]
    elif args.start_slide:
        pages = list(range(args.start_slide, n_pages + 1))
    else:
        pages = list(range(1, n_pages + 1))

    print(f"🎞  {args.pdf}  ({n_pages} slides)  bank={bank}  model={MODEL}")
    print(f"    Output → {out_path}  |  slides: {len(pages)}")

    if args.dry_run:
        print(f"\nDRY RUN — {len(pages)} slide(s) would be processed:")
        for p in pages:
            print(f"  slide {p:02d}")
        return

    client = genai.Client()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    used_names: set[str] = set(wb.sheetnames)
    index: list[dict] = []
    total_cost = 0.0

    for pg in pages:
        try:
            points, cost = process_slide(
                client, args.pdf, pg,
                bank, doc_title, doc_date,
                audit_dir,
                force=args.force,
            )
        except Exception as e:
            print(f"  slide {pg:02d}  ❌ {e}")
            continue

        total_cost += cost

        if not points:
            continue

        slide_title = points[0].slide_title or f"Slide {pg}"
        sname = tab_name(used_names, pg, slide_title)
        if sname in wb.sheetnames:
            del wb[sname]
        ws = wb.create_sheet(title=sname)

        render_slide(ws, points, pg, slide_title,
                     bank, doc_title, doc_date, brand)

        index.append({
            "slide_num":   pg,
            "slide_title": slide_title,
            "sheet":       sname,
            "n_points":    len(points),
            "has_chart":   any(p.source in ("chart", "unverified") for p in points),
        })

        build_contents(wb, index, bank, doc_title, doc_date, brand)
        wb.save(out_path)
        print(f"    💾 saved  |  running cost ≈ ${total_cost:.4f}")

    u = _run_usage
    print(f"\n✅ Done → {out_path}")
    print(f"   {u['calls']} API calls  "
          f"input={u['prompt']:,}  output={u['output']:,}  "
          f"≈ ${u['cost']:.4f}")


if __name__ == "__main__":
    main()
