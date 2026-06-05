# CLAUDE.md

Persistent context for Claude Code sessions on FinDocIQ. Read this first every session.

**Gemini is the runtime model the pipeline calls. Claude Code is the development agent that builds the pipeline.** Never conflate the two.

---

## What this project is

Two extraction pipelines for Singapore bank financial documents (DBS, OCBC, UOB):

1. **Pillar 3 pipeline** — extracts regulatory disclosure tables (Pillar 3, LCR, NSFR) from dense ~100-page PDFs into structured Excel workbooks. One tab per table.
2. **CFO Presentation pipeline** — extracts charts, waterfalls, P&L tables from 20–30 slide CFO decks into structured Excel workbooks. One tab per slide.

Both pipelines use Gemini Vision. The Pillar 3 pipeline sends native PDF bytes. The CFO pipeline sends PNG images.

---

## Folder structure (DELIVERABLE/)

```
DELIVERABLE/
  pillar3/
    PASS1_TOC.py              ← TOC extraction — zero API, pure Python + pdfplumber
    PASS2_Extract_to_Excel.py ← Gemini extraction → Excel (Gemini 2.5 Flash, PDF input)
    compare_excel.py          ← Validate output vs baseline

  CFO_Presentations/
    SLIDE_Extract.py          ← Full 4-pass pipeline (classify → extract → validate → render)
    chart_contracts.json      ← Per chart-type reading contracts (approved + pending_review)
    test_results/             ← E-12 single vs multi-pass test JSONs

  demo/
    generate_demo.py          ← Builds demo/index.html (the public-facing site)
    index.html                ← Self-contained demo site (regenerate with generate_demo.py)

  outputs/
    pillar3/                  ← dbs/ocbc/uob_pillar3.xlsx, toc.json, baselines, audit/
    CFO_Presentation/         ← dbs/ocbc/uob_slides.xlsx, run_summaries, audit/
```

---

## How to run

```bash
export GEMINI_API_KEY=...

# Pillar 3 — TOC (zero API, ~5s)
python3 DELIVERABLE/pillar3/PASS1_TOC.py "OCBC_4Q25_Pillar 3.pdf"

# Pillar 3 — Extract (~10-15 min, ~$0.50)
python3 DELIVERABLE/pillar3/PASS2_Extract_to_Excel.py "OCBC_4Q25_Pillar 3.pdf"

# CFO Slides — Extract all slides (~12-15 min, ~$0.18)
python3 DELIVERABLE/CFO_Presentations/SLIDE_Extract.py OCBC4Q25_CFO_presentation.pdf --force

# CFO Slides — Single slide re-run (preserves other tabs)
python3 DELIVERABLE/CFO_Presentations/SLIDE_Extract.py OCBC4Q25_CFO_presentation.pdf --slide 6 --force

# Rebuild demo site
python3 DELIVERABLE/demo/generate_demo.py --open
```

---

## Documentation sync rules — FOLLOW ON EVERY RELEVANT CHANGE

These are non-negotiable. When any of the following happen, update ALL relevant files before committing.

### On any workflow / architecture change:
| File | What to update |
|---|---|
| `MDs/DEVLOG.md` | Add E-XX entry: hypothesis, method, outcome, decision, learning |
| `MDs/TECHNICAL_DOCUMENTATION.md` | Update file map, script descriptions, flags, schema if changed |
| `CLAUDE.md` | Update Current state + Session log |
| `DELIVERABLE/demo/index.html` | Regenerate via `python3 demo/generate_demo.py` |

### On user command to "update the story" or "log the experiment":
Run all four updates above before pushing. Do not push partial updates.

### DEVLOG.md — when to add an entry:
- Any new extraction approach tried (even if it failed)
- Prompt changes made and tested
- Schema changes (add/remove fields)
- Architecture decisions (routing logic, pass structure)
- Renderer fixes or layout changes
- Benchmark results / accuracy comparisons

Each entry format:
```
### E-XX — Title ● STATUS / Date
| | |
|---|---|
| **Hypothesis** | ... |
| **Method** | ... |
| **Outcome** | ... |
| **Decision** | ... |
| **Learning** | ... |
```

### TECHNICAL_DOCUMENTATION.md — when to update:
- New script added or deleted → update file map table
- Script moved → update path
- New CLI flag → update usage section
- Output schema changes → update schema section
- Cost profile changes significantly → update cost section

---

## Design principles (never violate)

**CFO Presentations pipeline:**
- **No schema pressure during visual reading.** The extraction prompt does not constrain Gemini's visual reasoning. Chart contracts teach *method* not expected values.
- **Single-pass for visual slides.** Charts, waterfalls, donuts → one Gemini call, image present throughout. No intermediate text description.
- **Multi-pass for text tables only.** Pass 1 describes + pre-maps. Pass 2 transcribes text-only. No image in Pass 2.
- **Renderer is type-aware, not prompt-aware.** Waterfall → vertical list. KPI → two-column table. Pivot → wide format. Do not change the prompt to fix rendering bugs.
- **Contracts teach method, not values.** No hardcoded bank figures, no assumed ring order, no sample numbers in contracts.

**Pillar 3 pipeline:**
- **Deterministic code > LLM for predictable structure.** TOC extraction, page routing (single/multiple/spanning), Excel rendering — all fixed Python. Gemini only handles table content.
- **Remove fields the model should not fill.** Schema gaming (E-08): if a field exists, Gemini will use it to compress output. Remove fields that should not be filled.
- **Audit everything.** Every Gemini call saves prompt, PDF/image, raw response, parsed JSON, token counts. Resume is free — re-runs skip cached slides.

---

## Current state (2026-06-05)

**CFO Presentations pipeline (SLIDE_Extract.py):**
- Hybrid routing: Pass 0 classify → visual slides → single-pass, text tables → multi-pass
- Renderer rewritten to be type-aware: waterfall=vertical list, KPI=two-col, pivot=wide format
- Extra fields appended as trailing columns, never interleaved. Metadata keys (overlay_series) excluded.
- DBS (30 slides, ~$0.12), OCBC (21 slides, ~$0.17), UOB (22 slides, ~$0.18) all extracted
- DBS slides.xlsx only has slide 5 — full re-run needed
- chart_contracts.json: 9 approved contracts, no hardcoded values

**Pillar 3 pipeline (PASS1_TOC + PASS2_Extract_to_Excel):**
- Output renamed: `{bank}_pillar3.xlsx` (was `{bank}.xlsx`)
- OCBC and UOB extracted. DBS extracted but may have stale GCell-era code changes.
- Merged cell detection deferred (E-09/E-10). No span encoding in production schema.

**Demo site (demo/generate_demo.py):**
- Toggleable architecture diagram: CFO Slides vs Pillar 3 flows
- Results browser: 51 slides filterable by bank + chart type, expand to see datapoints
- Download buttons: Pillar 3 xlsx (dbs/ocbc/uob) + CFO slides xlsx
- PNGs not embedded (too large). Re-run generate_demo.py after new extractions.

**Folder layout:**
- `DELIVERABLE/slides/` → renamed to `DELIVERABLE/CFO_Presentations/`
- `outputs/slides/` → renamed to `outputs/CFO_Presentation/`
- `reports/` → merged into `demo/`
- Pillar 3 audit: `outputs/pillar3/audit/dbs|ocbc|uob/`
- Root `out/` directory: old working directory, duplicate of pillar3 audit — can be deleted

## Next tasks (priority order)

1. Re-run DBS CFO slides with `--force` (only slide 5 currently in Excel)
2. Re-run OCBC slide 12 with `--slide 12 --force` (failed on previous run)
3. Test renderer output on waterfall slides — verify vertical layout in Excel
4. Consider removing slide PNGs from index.html to reduce size (currently 27MB)
5. Overlay Series column rendering fix — stacked_bar_with_overlay CIR line
6. Implement deterministic pdfplumber x-position span pre-computation for OCBC Pillar 3 (E-09)

## Session log

> Newest at top. One entry per session.

- **2026-06-05** — Major session: folder restructure (slides→CFO_Presentations, outputs/slides→CFO_Presentation, reports→demo), hybrid routing architecture (single-pass visual / multi-pass text), prompt rewrites, renderer rewritten type-aware (waterfall vertical, KPI two-col, extra fields trailing), demo site built with toggleable architecture diagram + results browser + download buttons, DEVLOG E-11/E-12/E-13 added, Pillar 3 output renamed to _pillar3.xlsx, DBS/OCBC/UOB slides all extracted. CLAUDE.md rewritten.

- **2026-06-03** — Major DELIVERABLE cleanup. Reverted GCell to list[str] (E-07/E-08). Added 503 retry backoff, audit-based resume. Ran two-pass merge experiment (E-10), confirmed failed. OCBC and UOB Pillar 3 extractions verified.
