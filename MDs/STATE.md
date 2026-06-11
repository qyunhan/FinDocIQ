# STATE.md — Current State & Next Tasks
(volatile — rewrite freely; do not treat as history)

---

## Current state (2026-06-11)

### Pillar 3 pipeline — PASS2_Extract_to_Excel.py

**Model:** `gemini-3.5-flash`

**Schema (verified against code):**
- `GCell.cell_state`: `Literal["reported", "nil", "empty", "grey", "zero"]`
  with `@model_validator(mode="before")` migration hook — legacy state names
  (`"suppressed"` etc.) auto-migrate. `values` field is `list[GCell]`.
- `GRow.row_type`: `Literal["section_header", "data", "total", "sub_header", "note"]`
  with `@model_validator(mode="before")` migration hook for legacy string values.
- Pydantic emits JSON-schema `enum` arrays → Gemini structured output does
  constrained decoding on both fields.

**Checkpoint/cache split (P3-E-16):**
- Partial chunk writes → `parsed.partial.json` (never `parsed.json`)
- Final write: atomic `os.replace(tmp, parsed.json)`
- `meta.json` carries `prompt_hash` (SHA1[:8] of `_PROMPT`), `pages`, `partial`
- Cache-load path invalidates on: partial=True, doc mismatch, pages mismatch,
  hash mismatch (or absent hash = legacy cache)

**Post-extraction transforms (applied at all 5 consumption points):**
1. `split_date_blocks(t)` — splits tables with ≥2 date-period section_header
   rows into one GTable per period; fires only on shared-boundary pages
2. `fill_parents(t)` — overwrites GRow.parent deterministically from level-walk;
   assigns synthetic ids ("h1", "h2") to unnumbered rows referenced as parents
3. `drop_next_section_tables(tables, unit)` — drops tables whose title scores
   higher against `next_leaf.title` than against own leaf title; fires only when
   `next_leaf.start_page == unit.pages[-1]` (shared boundary page)

**Validation ladder (zero API cost):**
- `validate_spans` — column count per row
- `validate_numbers` — number multiset vs PDF text (with section_id noise suppression)
- `validate_labels` — duplicate stripped labels among data/total rows (row-shift detector)
- All three persisted to `meta.json["validation"]`; label_issues prints `⚠` but does not gate

**Outputs:**
- Excel: `DELIVERABLE/outputs/pillar3/{bank_slug}_pillar3.xlsx`
- Audit: `DELIVERABLE/outputs/pillar3/audit/{bank}/{doc_stem}/{unit_id}/`
  - `parsed.json` — complete extraction (GTable schema)
  - `meta.json` — provenance, usage, validation results

**OCBC 4Q25 extraction status:**
- 9.3: cached (gemini-2.5-flash, old prompt_hash — will auto-invalidate)
- 9.4: fresh (gemini-3.5-flash), 14 tables, 9 number_issues (boundary noise)
- 9.5: cached (gemini-2.5-flash, old prompt_hash — will auto-invalidate)
- 9.1, 9.2: no cache — need fresh extraction

### CFO Presentations pipeline — SLIDE_Extract.py
- Hybrid routing: Pass 0 classify → visual slides → single-pass; text tables → multi-pass
- DBS (30 slides), OCBC (21 slides), UOB (22 slides) all extracted
- chart_contracts.json: 9 approved contracts, no hardcoded values

### Demo site
- Toggleable architecture diagram, results browser, download buttons
- Regenerate with: `python3 DELIVERABLE/demo/generate_demo.py --open`

---

## Next tasks (priority order)

1. Run OCBC 4Q25 sections 9.1 + 9.2 (no cache — fresh extraction needed)
   - Tripwire: `grep -E "duplicate|⚠|✂|❌|invalidat|resumed" run.log`
   - Check 9.1 column headers — must be descriptive phrases, not `(a)(b)(c)`
2. Run 9.3 + 9.5 (will auto-invalidate on prompt_hash mismatch, re-extract)
   - Watch 9.5 tab titles — should NOT contain "Past Due Loans but Not Impaired"
     (that belongs to 9.4; if it appears, prev_leaf filter needed)
3. CLAUDE.md sidebar: add instruction to grep/jq audit files rather than reading whole
4. Wire `_is_text_only` into `validate_numbers` deficit loop to suppress MAS Notice
   637 false positive on DBS pages
5. Re-run audit sweep after 9.1–9.5 to confirm A.12.2.11 is the remaining DBS worklist

---

## Session log (newest first)

- **2026-06-11** — Implemented `split_date_blocks`, `fill_parents`, `_apply_transforms`,
  `drop_next_section_tables` with `next_leaf` in `build_units`. Added `validate_labels`
  row-shift detector. Switched model to `gemini-3.5-flash`. Rewrote CLAUDE.md; created
  STATE.md. All post-extraction transforms verified against OCBC 4Q25 9.4 cache.

- **2026-06-10** — Checkpoint/cache split (parsed.partial.json + atomic rename),
  prompt_hash cache invalidation, Literal enums on GCell.cell_state and GRow.row_type
  with model_validator(mode="before") migration hooks. All 4 acceptance cases pass.
  Audit sweep updated to pass section_ids. validate_numbers noise suppression (section-id
  dot-prefix chains, text-only large integers).

- **2026-06-05** — Major session: folder restructure, hybrid routing architecture,
  prompt rewrites, renderer rewritten type-aware, demo site built, DEVLOG E-11/E-12/E-13
  added, Pillar 3 output renamed to _pillar3.xlsx, DBS/OCBC/UOB slides all extracted.

- **2026-06-03** — Major DELIVERABLE cleanup. GCell span/merge schema tried and abandoned (E-07/E-08 legacy).
  503 retry backoff. Audit-based resume. Two-pass merge experiment (E-10) confirmed failed.
  OCBC and UOB Pillar 3 extractions verified.
