# RUN GUIDE — how to run FinDocIQ and see every output

This is the operator's manual. It tells you exactly where to put a PDF, what
command to run at each step, and which file to open to check the result. Every
step writes an inspectable artifact to `out/`.

---

## 0. One-time setup

```bash
cd findociq
pip install google-genai pypdfium2 openpyxl pillow pydantic
pip install docling          # optional: the high-fidelity table tool

export GEMINI_API_KEY=your_key_here   # runtime only — never commit or paste in chat
```

> **Put your PDF here:** copy it into the project root, e.g. `cp ~/Downloads/DBS_1Q26_Pillar3.pdf ./dbs.pdf`.
> Every command below takes the PDF path as its argument.

> **The most important rule:** if a command prints a `!! WARNING: no GEMINI_API_KEY`,
> it fell back to a crude deterministic mode that is *unreliable* for judgment
> steps (like the TOC). Set the key and re-run. The warning is the system being
> honest, not a bug.

---

## The workflow at a glance

| Step | Command | Open to check | Agentic? |
|------|---------|---------------|----------|
| 1+2 | `python phase1.py dbs.pdf` | `out/step1_toc.json`, `out/step2_table_map.csv` | TOC yes; counting no |
| 2b | `python phase2.py dbs.pdf` | `out/step2_docling/<id>.json` | no (docling) |
| 3 | `python extract_tables.py dbs.pdf` | `out/step3_extracted/<id>.json` | yes (the core) |
| 4 | `python render_all.py` | `out/step4_output.xlsx` | no (deterministic) |
| review | `python reviewer.py` | prints per-table flags | no |

---

## Step 1 + 2 — map the document

```bash
python phase1.py dbs.pdf
python phase2.py dbs.pdf
```

**What happens.**
- *Step 1 (agentic):* the agent reads the document and extracts the TOC — doc
  meta, section hierarchy, and each section's physical page range.
- *Step 2 (deterministic):* Python counts tables per page and classifies layout
  (none / single / multiple / spanning), flagging ambiguous pages.

**What to open.**
- `out/step1_toc.json` — **verify the content page is right.** Check every section
  is present with the correct page range and sub-sections. This is your manual
  gate before any extraction spends tokens.
- `out/step2_table_map.csv` — **open in Excel and eyeball.** One row per page:
  estimated table count, layout signals, and a `needs_agent_review` flag for
  pages where the count is uncertain. This is the human-checkable validation
  table you asked for.

> **Why counting is deterministic, not agentic:** once you have the page ranges,
> counting tables is structural — cheap Python. An agent re-checking correct
> Python output just burns tokens. Only the `needs_agent_review=True` rows are
> worth an agent's eyes; everything else is settled.

---

## Step 3 — extract each table (the agentic core)

```bash
python extract_tables.py dbs.pdf
```

**What happens, per table** (looping over the table map from step 2):
1. **Cut the PDF to that table's pages only.** The agent gets ONLY that table's
   context — no surrounding noise — which sharply improves accuracy.
2. **Choose the tool.** Born-digital page → text-first (cheap, reliable).
   Visual/merged/shaded table → render the image as a fallback. The agent
   decides based on what it sees; you don't hardcode it.
3. **Extract to the strict schema** (`FaithfulTable`) using the prompt in
   `prompts/extraction.py`.
4. **Review immediately.** A review check scores the table; on a discrepancy it
   triggers a prompt-iteration (see `prompts/self_iterate.py`) and re-extracts
   that one table.

**What to open.**
- `out/step3_extracted/<table_id>.json` — the structured extraction for each table.
- Run `python reviewer.py` to get per-table arithmetic flags.

---

## Step 4 — render Excel (deterministic, automatic)

```bash
python render_all.py
```

Produces:
- `out/step4_output.xlsx` — the faithful workbook: same rows, hierarchy columns,
  shaded cells, dashes, footnotes. **Open and human-verify.**

> **Why this is NOT the agent's job:** given the extracted representation, there
> is exactly one correct workbook. Making it agentic would destroy
> reproducibility and the audit story. The writer is fixed, inspectable code.

---

## Scoring & the refinement loop

To measure against hand-verified ground truth (currently the DBS SGD LCR, p12):

```bash
python measure_dbs.py          # writes out/dbs_accuracy.json
```

Read the headline `cell_accuracy_pct` and, most importantly, `wrong_values`
(drive this to 0 first). The full per-cell error list is in the JSON.

Then ask the critic for the next prompt improvement:

```bash
python run_critic.py           # prints a proposed edit + review checklist
```

Apply the approved edit to `prompts/extraction.py`, bump `PROMPT_VERSION`,
log the run, and re-measure. See `prompts/README.md` for the full loop and the
guardrails (never hardcode answers, never weaken grounding, human approves).

---

## Token efficiency (built in)

- **Step 2 is free** (deterministic) — no model call to count tables.
- **Step 3 cuts context to one table's pages** — the agent never reads the whole
  doc per table.
- **Text-first for born-digital** — image rendering (more tokens) only as a
  fallback.
- **Conditional dual-pass** (`cost_gate.py`) — the second verification pass runs
  only on flagged/high-stakes tables, ~40% cheaper than verifying everything.
- Measured cost on this DBS doc: well under 0.1 cent/page on a Flash-class model.

---

## If something looks wrong

- `out/step1_toc.json` has too many / wrong sections → you're on the deterministic
  fallback. Set `GEMINI_API_KEY` and re-run; the agent TOC is far better.
- A table flagged by `reviewer.py` → open its `out/step3_extracted/<id>.json`,
  compare to the PDF, and let the critic propose a prompt fix.
- Excel looks off but the JSON is right → that's a writer bug (deterministic),
  fix it in `writer.py`; do NOT push the fix into a prompt.
