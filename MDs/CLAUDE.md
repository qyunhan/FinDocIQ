# CLAUDE.md  (stable — edit rarely, keep under ~120 lines)

**Gemini is the runtime model the pipeline calls. Claude Code is the
development agent that builds it. Never conflate the two.**

## Source-of-truth hierarchy (read before trusting anything)
1. **The CODE is ground truth.** When any .md disagrees with the code, the
   code wins and the .md is the bug — fix the .md, not the code.
2. `MDs/STATE.md` — current state + next tasks (volatile, rewrite freely)
3. `MDs/DEVLOG.md` — append-only experiment history. NEVER edit old entries.
4. This file — identity, principles, commands (stable)

Do NOT act on "current state" claims anywhere else (session logs, chat
history, code comments). `MDs/TECHNICAL_DOCUMENTATION.md` is deprecated —
content folded into DEVLOG and STATE. Do not update it; treat as deleted.

## Sync rules (on any workflow / architecture / schema change)
| File | Action |
|---|---|
| `MDs/DEVLOG.md` | Append `P3-E-xx` / `CFO-E-xx` entry (never edit old) |
| `MDs/STATE.md` | Rewrite current state + session log entry |
| `FinDocIQ_DevLog_vN.docx` | Leadership rendering — regenerate per release, not per change |
| `demo/index.html` | Regenerate after new extractions |

## What this project is

Two extraction pipelines for Singapore bank financial documents (DBS, OCBC, UOB):

1. **Pillar 3 pipeline** — extracts regulatory disclosure tables from dense
   ~100-page PDFs into structured Excel workbooks. One tab per table.
2. **CFO Presentation pipeline** — extracts charts, waterfalls, P&L tables
   from 20–30 slide CFO decks. One tab per slide.

Both use Gemini Vision. Pillar 3 sends native PDF bytes; CFO sends PNG images.

## Folder structure

```
DELIVERABLE/
  pillar3/
    PASS1_TOC.py               ← TOC extraction — zero API, pure Python
    PASS2_Extract_to_Excel.py  ← Gemini extraction → Excel
  CFO_Presentations/
    SLIDE_Extract.py / chart_contracts.json
  demo/  generate_demo.py / index.html

DELIVERABLE/outputs/pillar3/
  {bank_slug}_pillar3.xlsx     ← shared across quarters (known collision
                                  hazard — per-document re-keying pending)
  {doc_stem}_toc.json
  audit/{bank}/{doc_stem}/{unit_id}/  parsed.json + meta.json
```

## How to run

```bash
export GEMINI_API_KEY=...
python3 DELIVERABLE/pillar3/PASS1_TOC.py "OCBC_4Q25_Pillar 3.pdf"
python3 DELIVERABLE/pillar3/PASS2_Extract_to_Excel.py "OCBC_4Q25_Pillar 3.pdf"

# Section re-run (cache resumes; --force to re-extract)
python3 DELIVERABLE/pillar3/PASS2_Extract_to_Excel.py "OCBC_4Q25_Pillar 3.pdf" \
    --section 9.4 2>&1 | tee run.log
grep -E "duplicate|⚠|✂|❌|invalidat|resumed" run.log
```

## Experiment numbering

One global sequence — prefixed `P3-E-xx` (Pillar 3) or `CFO-E-xx` (CFO).
Legacy unprefixed entries run E-01..E-15 and keep their numbers unchanged.
Prefixed sequence starts at 16. Next free number at top of `MDs/DEVLOG.md`.

## Verification norms (non-negotiable)

- Every task ends with a verification step the human can run.
- Before claiming a fix "didn't work", check whether the run resumed from
  cache (`grep "resumed" run.log`). Prompt/schema changes only affect fresh
  calls — cache-resume replays the old extraction unchanged.
- Never explain a validator result by reusing a prior-session diagnosis —
  read the current `parsed.json`/`meta.json` and cite it.
- Read audit files with `grep`/`jq`/`python3 -c` one-liners; do not `cat`
  large JSONs into context.

## Design principles (never violate)

**Both pipelines:**
- **Deterministic code > LLM for predictable structure.** TOC, routing,
  Excel rendering — fixed Python. Gemini only handles table/chart content.
- **Audit everything.** Every call saves parsed JSON + meta. Resume is free.
- **Schema validity ≠ semantic correctness** (P3-E-17). Pydantic validates
  structure; validators that read the PDF are needed for meaning.
- **Remove fields the model should not fill** (E-08, legacy). If a schema
  field exists, Gemini will use it to compress output. Remove it.

**Pillar 3 pipeline:**
- **Prompt is a prior; code is the contract** (P3-E-19). Gemini ignores
  prompt anchors under structured-output pressure — enforce deterministically.
- **Only ask the model for what requires perception; derive the rest.**
  Parents, date-block splits, boundary filtering — post-extraction Python.
- **Checkpoint and cache never share a filename** (P3-E-16). Partial writes
  go to `parsed.partial.json`; atomic rename produces `parsed.json`.

**CFO Presentations pipeline:**
- **No schema pressure during visual reading.** Contracts teach method only.
- **Single-pass for visual slides; multi-pass for text tables only.**
- **Renderer is type-aware, not prompt-aware.**
