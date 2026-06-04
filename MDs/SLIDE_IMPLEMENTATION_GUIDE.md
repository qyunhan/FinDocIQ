# SLIDE_Extract.py — Implementation Guide
## Two-Pass Visual Read + Chart Contracts

**For Claude Code. Read before touching any code.**

---

## What this guide covers

Two targeted changes to `DELIVERABLE/SLIDE_Extract.py`:

1. **Pass 0 — classify slide** (new file: `passes/pass0_classify.py`)
2. **Pass 1 prompt upgrade** — inject chart contracts + two-pass visual read instruction
3. **`chart_contracts.json`** — drop into project root

Everything else (Pass 2, validation, Pass 3, DataPoint schema, audit trail,
coerce, tab naming, cost tracking) stays exactly as-is. Do not touch them.

---

## Context: why these changes

**The gap in the current pipeline:**

Pass 1 describes the slide and Pass 2 extracts values — but both happen under
schema pressure, with Gemini doing visual reading and output formatting
simultaneously. This causes two failure modes:

1. **Sign errors on waterfalls** — Gemini reads the colour legend but applies
   signs at the same moment it's filling the schema, and gets it wrong on
   ambiguous bars.

2. **Misread values in dense charts** — Gemini estimates from bar height rather
   than reading the printed label, because it never explicitly inventories all
   visible numbers before assigning them to structure.

**The fix:**

Two additions that mirror how careful manual extraction works:

- **Chart contracts** teach Gemini the data model for each visual type *before*
  it starts reading values. Injected into Pass 1 so Gemini commits to the
  correct structure during the description phase, not during extraction.

- **Two-pass visual read instruction** forces Pass 1 to explicitly separate
  perceptual work (what numbers are printed?) from structural assignment (which
  number belongs where?). The inventory becomes a scratchpad that surfaces
  misreads before they propagate to Pass 2.

---

## Change 1 — Add `chart_contracts.json` to project root

Drop the file as-is into `DELIVERABLE/`. It is a living registry — `"status":
"approved"` entries are injected into Pass 1. Claude Code should never modify
the content of approved contracts without a documented reason.

**Location:** `DELIVERABLE/chart_contracts.json`

```json
{
  "waterfall": {
    "status": "approved",
    "contract": "..."
  },
  "stacked_bar_with_overlay": { ... },
  ...
}
```

(Full file provided separately as `chart_contracts.json`.)

---

## Change 2 — New file: `passes/pass0_classify.py`

Cheap micro-call. Identifies element types on the slide so Pass 1 knows which
contracts to inject. ~150 tokens total.

**Create this file exactly as specified. Do not add fields or change the return
type.**

```python
# passes/pass0_classify.py
import json
from google.genai import types as gtypes

KNOWN_TYPES = {
    "text_table", "waterfall", "stacked_bar", "stacked_bar_with_overlay",
    "trend_line", "kpi_grid", "pie", "donut_dual_ring",
    "npa_movement_table", "none"
}

CLASSIFY_PROMPT = """Look at this slide and list every distinct visual data element type present.

Return ONLY a JSON array of strings using these type names:
  "text_table"               - a printed table with rows and columns
  "waterfall"                - a bridge/waterfall chart showing running total deltas
  "stacked_bar"              - bars made of stacked coloured segments, no overlay line
  "stacked_bar_with_overlay" - stacked bars PLUS a trend line (%, bps) on the same axis
  "trend_line"               - line chart showing values over time, no bars
  "kpi_grid"                 - individual KPI metric boxes or callout figures
  "pie"                      - pie or donut chart (single ring)
  "donut_dual_ring"          - two concentric donut rings representing two time periods
  "npa_movement_table"       - NPA roll-forward table (opening + flows = closing)
  "none"                     - no data elements (title / agenda / closing slide)

If you see something that does not fit any of the above, invent a short
snake_case name for it and include it in the array.

Examples:
  ["text_table", "waterfall"]
  ["stacked_bar_with_overlay", "kpi_grid"]
  ["none"]
  ["text_table", "bullet_bridge"]   <- invented type
"""


def classify_slide(client, img_bytes: bytes, model: str) -> list[str]:
    """
    Pass 0: micro-call to identify element types on the slide.
    Returns list of type strings including any unknown/invented types.
    Cost: ~$0.0001 per slide at Flash rates.
    """
    img_part = gtypes.Part.from_bytes(data=img_bytes, mime_type="image/png")

    try:
        resp = client.models.generate_content(
            model=model,
            contents=[img_part, CLASSIFY_PROMPT],
            config=gtypes.GenerateContentConfig(temperature=0.0),
        )
    except Exception:
        return []  # fallback: Pass 1 gets all contracts

    raw = (resp.text or "").strip()

    # Strip markdown fences if present
    if raw.startswith("```"):
        raw = raw.split("```")[1].lstrip("json").strip()
        raw = raw.rsplit("```")[0].strip()

    try:
        result = json.loads(raw)
        if isinstance(result, list):
            return [str(t).strip() for t in result]
    except Exception:
        pass

    return []  # fallback: Pass 1 gets all contracts


def split_known_unknown(types_found: list[str]) -> tuple[list[str], list[str]]:
    """Split into known types (have approved contracts) and unknown types."""
    known   = [t for t in types_found if t in KNOWN_TYPES and t != "none"]
    unknown = [t for t in types_found if t not in KNOWN_TYPES and t != "none"]
    return known, unknown
```

---

## Change 3 — Modify `passes/pass1_describe.py`

### 3a — Add contract loader at top of file

```python
# passes/pass1_describe.py  — ADD these imports and functions at the top

import json
import os
import re


def load_contracts(contracts_path: str = "chart_contracts.json") -> dict[str, str]:
    """Load approved chart contracts from the registry."""
    if not os.path.exists(contracts_path):
        return {}
    with open(contracts_path) as f:
        raw = json.load(f)
    return {
        k: v["contract"]
        for k, v in raw.items()
        if v.get("status") == "approved"
    }


def build_contracts_block(known_types: list[str], contracts: dict[str, str]) -> str:
    """Inject contracts for the element types present on this slide."""
    if not known_types:
        return ""

    blocks = []
    for t in known_types:
        if t in contracts:
            blocks.append(f"--- CONTRACT: {t.upper()} ---\n{contracts[t]}")

    if not blocks:
        return ""

    return (
        "\n\nCHART READING CONTRACTS FOR THIS SLIDE:\n"
        "Apply these before reasoning about any chart element.\n\n"
        + "\n\n".join(blocks)
    )


UNKNOWN_TYPE_TEMPLATE = """
--- CONTRACT: {unknown_type} (UNKNOWN — DERIVE YOUR OWN) ---
You identified "{unknown_type}" which has no predefined reading contract.

Before extracting, derive a contract for it:

1. DATA MODEL: what underlying data structure generated this visual?
   (what would the raw spreadsheet rows look like?)

2. ARITHMETIC CONSTRAINT: is there a summation relationship between values?
   If none, state "no constraint".

3. HOW TO READ: what visual signals encode the values?
   (position, height, colour, label, size, angle)

4. SIGN RULE: can values be negative?
   If so, what visual signal indicates sign?

Write this contract explicitly in your description under the heading:
  DERIVED CONTRACT: {unknown_type}
"""


def build_unknown_contracts_block(unknown_types: list[str]) -> str:
    if not unknown_types:
        return ""
    return "\n\n".join(
        UNKNOWN_TYPE_TEMPLATE.format(unknown_type=t)
        for t in unknown_types
    )


def save_derived_contract(description: str,
                          unknown_types: list[str],
                          contracts_path: str = "chart_contracts.json"):
    """
    Parse any derived contracts from the Pass 1 description and save
    them to chart_contracts.json with status=pending_review.
    """
    if not unknown_types:
        return

    existing = {}
    if os.path.exists(contracts_path):
        with open(contracts_path) as f:
            existing = json.load(f)

    changed = False
    for t in unknown_types:
        pattern = rf"DERIVED CONTRACT:\s*{re.escape(t)}\s*\n(.*?)(?=\n---|\Z)"
        match = re.search(pattern, description, re.DOTALL | re.IGNORECASE)
        if match and t not in existing:
            existing[t] = {
                "status":   "pending_review",
                "contract": match.group(1).strip()
            }
            print(f"  📝 new contract derived for '{t}' — saved as pending_review")
            changed = True

    if changed:
        with open(contracts_path, "w") as f:
            json.dump(existing, f, indent=2)
```

### 3b — Replace the PASS1_PROMPT constant

**Find this in the current file:**
```python
PASS1_PROMPT = """Examine this bank CFO presentation slide carefully.
...
Do not extract values yet."""
```

**Replace it with this:**
```python
PASS1_PROMPT = """Examine this bank CFO presentation slide carefully.

═══════════════════════════════════════════════════
STEP 1 — INVENTORY (do this first, before any structure)
═══════════════════════════════════════════════════

Scan the entire slide and list EVERY printed number you can see.
For each number write:
  - the number exactly as printed (e.g. "2,296", "(1,236)", "1.91%")
  - its approximate location on the slide (e.g. "top-left bar", "row 3 col 2")
  - any adjacent label text (e.g. "Net Interest Income", "Volume", "4Q25")

Do not skip any number, even if you are unsure what it represents.
If a number is partially obscured or hard to read, write it with a "?" flag.

═══════════════════════════════════════════════════
STEP 2 — STRUCTURE
═══════════════════════════════════════════════════

For each distinct data element on the slide, describe:

1. TYPE: what kind of visual is it?
   (text_table | waterfall | stacked_bar | stacked_bar_with_overlay |
    trend_line | kpi_grid | pie | donut_dual_ring | npa_movement_table | other)

2. TITLE: the label printed above it (verbatim)

3. STRUCTURE:
   - text_table / npa_movement_table: how many rows, how many columns,
     what are the column headers
   - waterfall: how many bars total, opening bar label, closing bar label,
     what does the colour legend say (e.g. "green = positive, red = negative"),
     are % labels printed on bars and what do they represent (YoY? QoQ?)
   - stacked_bar / stacked_bar_with_overlay: how many time periods,
     how many stack components, what are the period labels, what are the
     component labels; if overlay line present — what is its name and unit
   - trend_line: how many series, how many periods, what are the series names
   - kpi_grid: how many KPIs, what are the labels
   - pie / donut_dual_ring: how many segments; for dual ring — which ring
     is the earlier period and which is the later period

4. UNITS: what unit are values in (S$m, S$b, %, bps, etc.)

5. VISUAL CONVENTIONS on this specific slide:
   - bold rows = totals?
   - indented rows = sub-items?
   - shaded/grey cells = not applicable?
   - bracket groupings?
   - any footnotes that change interpretation?

═══════════════════════════════════════════════════
STEP 3 — ASSIGN AND VERIFY
═══════════════════════════════════════════════════

Using your inventory from Step 1 and the structure from Step 2:

For each element, assign every number from your inventory to its structural
role (which series, which period, which component).

Then apply the arithmetic constraint from the chart contract (if one exists):
  - Waterfall: start + sum(signed deltas) = end. If it does not balance,
    find the misread or wrong sign in your inventory before proceeding.
  - Stacked bar: components sum to bar total. Check each period.
  - NPA table: opening + flows = closing. Check each column.
  - Pie / donut: segments sum to ~100% (or printed total).

If the arithmetic fails: DO NOT proceed. Return to your inventory, find the
error (wrong number, wrong sign, missed bar), correct it, and re-verify.

Write out the arithmetic check explicitly:
  e.g. "Waterfall check: 9,755 + 658 - 1,236 - 27 = 9,150 ✓"
  e.g. "Stacked bar 4Q25: 969 + 334 + 256 = 1,559 ✓"

Only after the arithmetic check passes should you summarise the element.

═══════════════════════════════════════════════════

Be specific and complete. This description — including the inventory and
arithmetic checks — is the input to the extraction step. Do not extract
structured values yet, but do write out all three steps above.
"""
```

### 3c — Update `describe_slide` function signature

**Find the current function:**
```python
def describe_slide(client, img_bytes: bytes, model: str) -> tuple[str, float]:
```

**Replace with:**
```python
def describe_slide(client, img_bytes: bytes, model: str,
                   known_types: list[str] | None = None,
                   unknown_types: list[str] | None = None) -> tuple[str, float]:
    """
    Pass 1: describe slide with two-pass visual read.
    Injects approved chart contracts for known_types.
    Asks Gemini to derive contracts for unknown_types.
    """
    contracts     = load_contracts()
    known_block   = build_contracts_block(known_types or [], contracts)
    unknown_block = build_unknown_contracts_block(unknown_types or [])

    prompt = PASS1_PROMPT + known_block + unknown_block

    # ... rest of existing function body unchanged ...
    # (the call to client.models.generate_content, cost logging, return)
```

---

## Change 4 — Modify `extract.py` (main entry point)

### 4a — Add imports at top

```python
from passes.pass0_classify import classify_slide, split_known_unknown
from passes.pass1_describe import save_derived_contract
```

### 4b — Replace `process_slide` function

The existing `process_slide` function needs Pass 0 inserted before Pass 1.
Find the function and modify it as follows.

**Current flow:**
```
render_page → Pass 1 describe → Pass 2 extract → validate → correct → save
```

**New flow:**
```
render_page → Pass 0 classify → Pass 1 describe (with contracts) → save derived contracts
           → Pass 2 extract → validate → correct → save
```

**Exact changes to `process_slide`:**

After `img_bytes = render_page(...)` and before Pass 1, insert:

```python
    # ── Pass 0: Classify ─────────────────────────────────────────
    types_path = os.path.join(audit_path, "element_types.json")

    if not force and os.path.exists(types_path):
        # Resume Pass 0 from audit
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
```

Then change the Pass 1 call from:
```python
    description, cost1 = describe_slide(client, img_bytes, MODEL)
```

To:
```python
    # ── Pass 1: Describe with contracts ──────────────────────────
    description, cost1 = describe_slide(
        client, img_bytes, MODEL,
        known_types=known_types,
        unknown_types=unknown_types,
    )
```

Then after saving `description.txt`, add:
```python
    # Save any derived contracts Gemini wrote for unknown types
    save_derived_contract(description, unknown_types)
```

Everything else in `process_slide` is unchanged.

---

## Change 5 — Audit trail additions

The audit folder for each slide now contains one new file:

```
outputs/audit/slides/{bank}_{doc}/
  slide_{N}/
    element_types.json    ← NEW: Pass 0 output  ["waterfall", "text_table"]
    description.txt       ← Pass 1 output (now includes inventory + contracts)
    datapoints.json       ← Pass 2 output
    meta.json             ← cost, tokens, validation errors
```

`element_types.json` is already written in the Pass 0 block above. No
additional code needed.

---

## What NOT to change

Do not touch any of the following:

- `passes/pass2_extract.py` — unchanged
- `passes/pass3_render.py` — unchanged
- `validation.py` — unchanged
- `models.py` — DataPoint schema unchanged
- `bank_config.py` — unchanged
- `utils.py` — unchanged
- The `self_check` field in Pass 2 prompt — unchanged
- The correction pass logic — unchanged
- Cost logging and audit save logic — unchanged

---

## Verification checklist (run after implementing)

Run a single known slide with `--slide N --force` and check:

```
outputs/audit/slides/{bank}/slide_{N}/element_types.json
```
Should contain a list like `["waterfall", "text_table"]` — not empty, not `["none"]`.

```
outputs/audit/slides/{bank}/slide_{N}/description.txt
```
Should contain three clearly labelled sections:
- **STEP 1** — a numbered list of every visible number with location
- **STEP 2** — structural description per element
- **STEP 3** — explicit arithmetic checks like "9,755 + 658 - 1,236 - 27 = 9,150 ✓"

If STEP 3 is absent or the arithmetic check is missing, the Pass 1 prompt was
not injected correctly.

```
outputs/audit/slides/{bank}/slide_{N}/datapoints.json
```
For a waterfall slide: check that `sign` fields are present on all bridge
components and that `start + sum(signed values) ≈ end`.

---

## Cost impact

Pass 0 adds ~$0.0001 per slide (150 tokens at Flash rates). Negligible.
Pass 1 grows by ~200-400 tokens from the contract injection. Still under
$0.01 per slide target.

The main benefit is not cost but accuracy: the inventory step surfaces
misreads before they reach Pass 2, and the arithmetic check in Pass 1
catches sign errors before the correction pass has to fire.

---

## Session log entry (add to DEVLOG.md)

```
### E-11 — Two-Pass Visual Read + Chart Contracts ● IMPLEMENTED

Hypothesis: Separating perceptual work (number inventory) from structural
assignment (which number belongs where) in Pass 1 reduces sign errors
on waterfalls and misread values in dense charts.

Method:
- Added chart_contracts.json registry with 9 approved types
- Added Pass 0 (classify_slide) — cheap micro-call to identify element types
- Modified Pass 1 prompt to enforce explicit three-step process:
  Step 1: inventory all visible numbers with location
  Step 2: describe structure per element
  Step 3: assign numbers to structure + verify arithmetic before proceeding

Arithmetic check in Pass 1 means sign errors are caught during description,
not after extraction. The inventory scratchpad surfaces misreads before
they propagate to Pass 2.
```
"""
