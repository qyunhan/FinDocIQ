# FinDocIQ Pipeline — Spec Delta
## Changes to PIPELINE_SPEC.md

---

## 1. New file: `chart_contracts.json`

Create this file in the project root. It is the living registry of
chart reading contracts. Human-reviewed entries are `"status": "approved"`.
Gemini-derived entries pending review are `"status": "pending_review"`.

```json
{
  "waterfall": {
    "status": "approved",
    "contract": "DATA MODEL: start_value + sum(signed_deltas) = end_value\nHOW TO READ:\n  - Each bar represents one delta in the running total\n  - Bar goes UP from running total → positive delta\n  - Bar goes DOWN from running total → negative delta\n  - If legend present: use it to confirm sign interpretation\n  - If no legend: bar direction relative to running total IS the sign\nCONSTRAINT: sum all deltas, verify start + sum = end\n  If arithmetic fails, your signs are wrong — flip bars until it balances\nEXTRACT AS: one DataPoint per bar, value must be signed"
  },
  "stacked_bar": {
    "status": "approved",
    "contract": "DATA MODEL: for each period, components stack bottom-to-top summing to total\nHOW TO READ:\n  - Read each segment value from its label\n  - Bottom segment = first component, top = last\n  - Total = sum of all segments for that period\n  - Period axis = x-axis labels\nCONSTRAINT: sum of components per period = printed total above bar\nEXTRACT AS: one DataPoint per (component, period) combination"
  },
  "trend_line": {
    "status": "approved",
    "contract": "DATA MODEL: for each series, one value per time period\nHOW TO READ:\n  - Each line = one series, identified by label at line end or in legend\n  - Each point value is printed above or below the point\n  - X-axis = periods, Y-axis = values\n  - Multiple lines = multiple series, extract all\nCONSTRAINT: none\nEXTRACT AS: one DataPoint per (series, period)"
  },
  "pie": {
    "status": "approved",
    "contract": "DATA MODEL: segments summing to total (100% or absolute)\nHOW TO READ:\n  - Each segment has label and value\n  - If % shown: should sum to ~100%\n  - If absolute: should sum to printed total\n  - Centre label if present = grand total\nCONSTRAINT: sum of segments ≈ 100% or centre total\nEXTRACT AS: one DataPoint per segment"
  },
  "kpi_grid": {
    "status": "approved",
    "contract": "DATA MODEL: individual metrics, no summation relationship\nHOW TO READ:\n  - Each KPI has label, current value, and often change vs prior period\n  - Change may be absolute, %, or both\nCONSTRAINT: none\nEXTRACT AS: one DataPoint per KPI, extra_fields for change values"
  },
  "text_table": {
    "status": "approved",
    "contract": "DATA MODEL: rows and columns, bold rows = totals, indented = sub-items\nHOW TO READ:\n  - Column headers define the value dimensions\n  - Bold label = total or subtotal row\n  - Indented label = sub-item of row above\n  - Grey/shaded cell = not applicable, value = empty string\n  - Parentheses = negative value, keep verbatim\nCONSTRAINT: sub-items should sum to their parent total\nEXTRACT AS: one DataPoint per (row, column) combination"
  }
}
```

---

## 2. New file: `passes/pass0_classify.py`

Insert a new Pass 0 before Pass 1. Cheap micro-call that returns
a list of element types on the slide. Used to select which contracts
to inject into Pass 1.

```python
import json
from typing import Any

KNOWN_TYPES = {
    "text_table", "waterfall", "stacked_bar",
    "trend_line", "kpi_grid", "pie", "none"
}

CLASSIFY_PROMPT = """
Look at this slide and list every distinct visual data element type present.

Return ONLY a JSON array of strings using these type names:
  "text_table"   - a printed table with rows and columns
  "waterfall"    - a bridge/waterfall chart showing running total deltas
  "stacked_bar"  - bars made of stacked coloured segments
  "trend_line"   - line chart showing values over time periods
  "kpi_grid"     - individual KPI metric boxes or callout figures
  "pie"          - pie or donut chart
  "none"         - no data elements (title/agenda/closing slide)

If you see something that doesn't fit any of the above, invent a
short snake_case name for it and include it in the array.

Examples:
  ["text_table", "waterfall"]
  ["stacked_bar", "trend_line", "kpi_grid"]
  ["none"]
  ["text_table", "waterfall", "bullet_bridge"]   ← invented type
"""


def classify_slide(client, img_bytes: bytes, model: str) -> list[str]:
    """
    Pass 0: micro-call to identify element types on the slide.
    Returns list of type strings including any unknown/invented types.
    """
    from google.genai import types as gtypes

    img_part = gtypes.Part.from_bytes(
        data=img_bytes,
        mime_type="image/png"
    )

    resp = client.models.generate_content(
        model=model,
        contents=[img_part, CLASSIFY_PROMPT],
        config=gtypes.GenerateContentConfig(temperature=0.0)
        # No response_mime_type — plain text, parse manually
    )

    raw = (resp.text or "").strip()

    # Strip markdown fences if present
    if raw.startswith("```"):
        raw = raw.split("```")[1].lstrip("json").strip()
        raw = raw.rsplit("```")[0].strip()

    try:
        types_found = json.loads(raw)
        if isinstance(types_found, list):
            return [str(t).strip() for t in types_found]
    except Exception:
        pass

    # Fallback: return empty means Pass 1 gets all contracts
    return []


def split_known_unknown(types_found: list[str]) -> tuple[list[str], list[str]]:
    known   = [t for t in types_found if t in KNOWN_TYPES]
    unknown = [t for t in types_found if t not in KNOWN_TYPES and t != "none"]
    return known, unknown
```

---

## 3. Modified file: `passes/pass1_describe.py`

### 3a. Add contract loader

```python
import json, os

def load_contracts(chart_contracts_path: str = "chart_contracts.json") -> dict:
    if not os.path.exists(chart_contracts_path):
        return {}
    with open(chart_contracts_path) as f:
        raw = json.load(f)
    # Only return approved contracts
    return {
        k: v["contract"]
        for k, v in raw.items()
        if v.get("status") == "approved"
    }


def build_contracts_block(known_types: list[str],
                           contracts: dict) -> str:
    """
    Inject only the contracts relevant to this slide.
    """
    if not known_types:
        return ""

    blocks = []
    for t in known_types:
        if t in contracts:
            blocks.append(
                f"--- CONTRACT: {t.upper()} ---\n{contracts[t]}"
            )

    if not blocks:
        return ""

    return (
        "\n\nCHART READING CONTRACTS FOR THIS SLIDE:\n"
        "Apply these when reasoning about each chart element.\n\n"
        + "\n\n".join(blocks)
    )
```

### 3b. Add unknown type instruction

```python
UNKNOWN_TYPE_TEMPLATE = """
--- CONTRACT: {unknown_type} (UNKNOWN — DERIVE YOUR OWN) ---
You identified "{unknown_type}" which has no predefined reading contract.

Before extracting, derive a contract for it:

1. DATA MODEL: what underlying data structure generated this visual?
   (what would the raw spreadsheet data look like?)

2. ARITHMETIC CONSTRAINT: is there a summation relationship between values?
   If none, state "no constraint".

3. HOW TO READ: what visual signals encode the values?
   (position, height, colour, label, size, angle)

4. SIGN RULE: can values be negative?
   If so, what visual signal indicates sign?

Write this contract explicitly in your description under the heading:
  DERIVED CONTRACT: {unknown_type}

This derived contract will be saved for future use and human review.
"""


def build_unknown_contracts_block(unknown_types: list[str]) -> str:
    if not unknown_types:
        return ""
    blocks = [
        UNKNOWN_TYPE_TEMPLATE.format(unknown_type=t)
        for t in unknown_types
    ]
    return "\n\n".join(blocks)
```

### 3c. Update `describe_slide` signature

```python
def describe_slide(client, img_bytes: bytes, model: str,
                   known_types: list[str] = None,
                   unknown_types: list[str] = None) -> tuple[str, float]:
    """
    Pass 1: describe slide and commit to chart contracts.
    known_types:   inject approved contracts for these types
    unknown_types: ask Gemini to derive its own contract for these
    """
    contracts     = load_contracts()
    known_block   = build_contracts_block(known_types or [], contracts)
    unknown_block = build_unknown_contracts_block(unknown_types or [])

    prompt = PASS1_PROMPT + known_block + unknown_block

    # ... rest of existing describe_slide call unchanged ...
```

---

## 4. New function: `save_derived_contract` in `passes/pass1_describe.py`

After Pass 1 completes, parse any derived contracts from the description
and save them to `chart_contracts.json` for human review.

```python
import re

def save_derived_contract(description: str,
                          unknown_types: list[str],
                          contracts_path: str = "chart_contracts.json"):
    """
    Parse derived contracts from Pass 1 description text.
    Save to chart_contracts.json with status=pending_review.
    """
    if not unknown_types:
        return

    existing = {}
    if os.path.exists(contracts_path):
        with open(contracts_path) as f:
            existing = json.load(f)

    for t in unknown_types:
        # Look for DERIVED CONTRACT: {t} section in description
        pattern = rf"DERIVED CONTRACT:\s*{re.escape(t)}\s*\n(.*?)(?=\n---|\Z)"
        match = re.search(pattern, description, re.DOTALL | re.IGNORECASE)

        if match:
            derived_text = match.group(1).strip()
            if t not in existing:
                existing[t] = {
                    "status":  "pending_review",
                    "contract": derived_text
                }
                print(f"  📝 New contract derived for '{t}' — "
                      f"saved as pending_review in {contracts_path}")

    with open(contracts_path, "w") as f:
        json.dump(existing, f, indent=2)
```

---

## 5. Modified file: `extract.py`

### Update `process_slide` to include Pass 0

```python
def process_slide(client, pdf_path, page_num, bank, doc_title,
                  doc_date, audit_dir, force=False):

    label       = f"slide_{page_num:02d}"
    audit_path  = os.path.join(audit_dir, label)
    os.makedirs(audit_path, exist_ok=True)

    # Paths
    types_path  = os.path.join(audit_path, "element_types.json")
    desc_path   = os.path.join(audit_path, "description.txt")
    dp_path     = os.path.join(audit_path, "datapoints.json")
    meta_path   = os.path.join(audit_path, "meta.json")

    # Resume from audit if available
    if not force and os.path.exists(dp_path):
        with open(dp_path) as f:
            raw = json.load(f)
        points = [DataPoint(**p) for p in raw]
        print(f"  slide {page_num:02d}  resumed ({len(points)} points)")
        return points, 0.0

    img_bytes  = render_page(pdf_path, page_num, scale=3.0)
    total_cost = 0.0

    # ── Pass 0: Classify ──────────────────────────────────────────
    if not force and os.path.exists(types_path):
        with open(types_path) as f:
            types_found = json.load(f)
        print(f"  slide {page_num:02d}  types resumed: {types_found}")
    else:
        types_found = classify_slide(client, img_bytes, MODEL)
        with open(types_path, "w") as f:
            json.dump(types_found, f)

    known_types, unknown_types = split_known_unknown(types_found)

    # Skip slides with no data elements
    if types_found == ["none"] or not types_found:
        print(f"  slide {page_num:02d}  — no data elements, skipping")
        return [], 0.0

    # ── Pass 1: Describe with targeted contracts ──────────────────
    description, cost1 = describe_slide(
        client, img_bytes, MODEL,
        known_types=known_types,
        unknown_types=unknown_types
    )
    total_cost += cost1

    with open(desc_path, "w") as f:
        f.write(description)

    # Save any derived contracts from unknown types
    save_derived_contract(description, unknown_types)

    # ── Pass 2: Extract ───────────────────────────────────────────
    # ... unchanged from existing spec ...
```

---

## 6. Audit trail additions

Each slide's audit folder now contains:

```
audit/{bank}_{doc}/slide_{N}/
  ├── element_types.json    ← NEW: Pass 0 output ["waterfall", "text_table"]
  ├── description.txt       ← Pass 1 output (now includes chart contracts)
  ├── datapoints.json       ← Pass 2 output
  └── meta.json             ← cost, tokens, validation errors
```

---

## 7. Summary of new call sequence

```
Pass 0 — classify_slide()
  ~150 tokens total, returns ["waterfall", "text_table"] etc.
  Saves element_types.json
  Splits into known_types + unknown_types

Pass 1 — describe_slide(known_types, unknown_types)
  Injects approved contracts for known_types only
  Asks Gemini to derive contract for unknown_types
  Saves description.txt + updates chart_contracts.json if new type found

Pass 2 — extract_slide(description)
  Unchanged — still injects description into prompt
  DataPoints now more accurate because Pass 1 committed to contracts

Validation — validate_self_check()
  Unchanged — generic arithmetic string parser

Pass 3 — render_to_excel()
  Unchanged
```

---

## 8. No changes needed to

- `models.py` — DataPoint schema unchanged
- `validation.py` — generic self_check unchanged  
- `passes/pass2_extract.py` — unchanged
- `passes/pass3_render.py` — unchanged
- `bank_config.py` — unchanged
