"""
generate_demo.py — FinDocIQ demo site.
Builds a self-contained index.html in the same folder.

Usage:
    cd DELIVERABLE/demo
    python3 generate_demo.py          # build index.html
    python3 generate_demo.py --open   # build + open in browser
"""
from __future__ import annotations
import re, json, base64, argparse, webbrowser
from datetime import datetime
from pathlib import Path

DEMO_DIR    = Path(__file__).parent.resolve()
DELIVERABLE = DEMO_DIR.parent
MDS_DIR     = DELIVERABLE.parent / "MDs"
OUTPUTS_DIR = DELIVERABLE / "outputs" / "CFO_Presentation"
AUDIT_DIR   = OUTPUTS_DIR / "audit"
CONTRACTS_F = DELIVERABLE / "slides" / "chart_contracts.json"

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def read_text(p) -> str:
    p = Path(p)
    return p.read_text(encoding="utf-8") if p.exists() else ""

def b64_img(p) -> str | None:
    p = Path(p)
    if not p.exists(): return None
    ext = p.suffix.lower().lstrip(".")
    mime = {"png":"image/png","jpg":"image/jpeg","jpeg":"image/jpeg"}.get(ext,"image/png")
    return f"data:{mime};base64,{base64.b64encode(p.read_bytes()).decode()}"

def b64_xlsx(p) -> str | None:
    p = Path(p)
    if not p.exists(): return None
    return base64.b64encode(p.read_bytes()).decode()

def jload(p):
    p = Path(p)
    if not p.exists(): return None
    try: return json.loads(p.read_text())
    except: return None

def esc(s): return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------

def load_all_slides() -> list[dict]:
    slides = []
    if not AUDIT_DIR.exists(): return slides
    for bank_dir in sorted(AUDIT_DIR.iterdir()):
        if not bank_dir.is_dir(): continue
        bank = bank_dir.name.split("_")[0].upper()
        doc  = "_".join(bank_dir.name.split("_")[1:])
        for sd in sorted(bank_dir.iterdir()):
            m = re.match(r"slide_(\d+)$", sd.name)
            if not m: continue
            num   = int(m.group(1))
            types = jload(sd/"element_types.json") or []
            if types == ["none"] or not types: continue
            meta  = jload(sd/"meta.json") or {}
            dps   = jload(sd/"datapoints.json") or []
            if not dps: continue
            title = dps[0].get("slide_title","") if dps else ""
            usages = meta.get("usages",[])
            branch = "single-pass" if any(str(u.get("pass",""))=="1s" for u in usages) else "multi-pass"
            img    = b64_img(sd/f"slide_{num:02d}.png") or b64_img(sd/"slide.png")
            slides.append({
                "bank": bank, "doc": doc, "num": num,
                "title": title or f"Slide {num}",
                "types": types, "n_pts": meta.get("n_points", len(dps)),
                "cost": meta.get("cost_usd", 0.0),
                "errors": meta.get("validation_errors") or [],
                "branch": branch, "image_b64": img,
                "datapoints": dps[:60],  # cap for HTML size
            })
    return slides

def load_run_summaries() -> list[dict]:
    out = []
    for f in sorted(OUTPUTS_DIR.glob("*_run_summary.json")):
        s = jload(f)
        if s: out.append(s)
    # also check for xlsx without summary
    for xl in OUTPUTS_DIR.glob("*_slides.xlsx"):
        bank = xl.stem.split("_")[0].upper()
        if not any(s.get("bank","").upper()==bank for s in out):
            out.append({"bank": bank, "document": xl.name,
                        "slides_processed":"—","api_calls":"—",
                        "est_cost_usd":0,"elapsed_human":"—","generated_at":""})
    return out

def load_contracts() -> dict:
    return jload(CONTRACTS_F) or {}

def load_devlog_experiments() -> list[dict]:
    md = read_text(MDS_DIR/"DEVLOG.md")
    if not md: return []
    exp_re = re.compile(
        r"###\s+(E-\d+)\s+[—–-]+\s+(.+?)\s+[●•]\s+(\w[\w\s/]+?)(?:\s*/[^\n]*)?\n",
        re.IGNORECASE)
    matches = list(exp_re.finditer(md))
    exps = []
    for i, m in enumerate(matches):
        eid = m.group(1)
        title = m.group(2).strip()
        st_raw = m.group(3).strip().upper()
        if "ADOPT" in st_raw: status = "adopted"
        elif "PARTIAL" in st_raw: status = "partial"
        elif "INSIGHT" in st_raw or "DESIGN" in st_raw: status = "insight"
        elif "IMPLEMENT" in st_raw: status = "implemented"
        else: status = "failed"
        body_start = m.end()
        body_end = matches[i+1].start() if i+1<len(matches) else len(md)
        body = md[body_start:body_end]
        def field(f):
            mm = re.search(rf"\|\s*\*\*{f}\*\*\s*\|\s*(.+?)(?=\n\||\Z)", body, re.DOTALL)
            return mm.group(1).strip()[:400] if mm else ""
        exps.append({"id":eid,"title":title,"status":status,
                     "outcome":field("Outcome"),"learning":field("Learning")})
    return exps

# ---------------------------------------------------------------------------
# HTML COMPONENTS
# ---------------------------------------------------------------------------

TYPE_COLOURS = {
    "waterfall":"#c0392b","stacked_bar":"#2980b9",
    "stacked_bar_with_overlay":"#1a6fa8","trend_line":"#27ae60",
    "kpi_grid":"#8e44ad","text_table":"#34495e","pie":"#d68910",
    "donut_dual_ring":"#e67e22","npa_movement_table":"#16a085","bar_chart":"#2980b9",
}
STATUS_CFG = {
    "failed":("#e74c3c","FAILED"),"adopted":("#27ae60","ADOPTED"),
    "partial":("#f39c12","PARTIAL"),"insight":("#8e44ad","INSIGHT"),
    "implemented":("#2980b9","IMPLEMENTED"),
}

def chip(t):
    bg = TYPE_COLOURS.get(t,"#7f8c8d")
    return f'<span class="chip" style="background:{bg}">{t.replace("_"," ")}</span>'

def badge(status):
    bg, lb = STATUS_CFG.get(status,("#7f8c8d",status.upper()))
    return f'<span class="badge" style="background:{bg}">{lb}</span>'

def bank_dot(bank):
    c = {"DBS":"#cc0000","OCBC":"#cc0000","UOB":"#1b6ec2"}.get(bank,"#888")
    return f'<span class="bank-dot" style="background:{c}">{bank}</span>'

# ---------------------------------------------------------------------------
# SECTION BUILDERS
# ---------------------------------------------------------------------------

def build_hero(slides, summaries):
    total_pts   = sum(s["n_pts"] for s in slides)
    total_slides = len(slides)
    banks = sorted({s["bank"] for s in slides})
    total_cost  = sum(s["cost"] for s in slides)  # sum per-slide costs from meta.json
    # count chart types
    type_counts = {}
    for s in slides:
        for t in s["types"]:
            type_counts[t] = type_counts.get(t,0)+1
    top_types = sorted(type_counts, key=lambda x:-type_counts[x])[:4]
    type_pills = "".join(f'<span class="hero-pill">{t.replace("_"," ")}</span>' for t in top_types)

    return f"""
    <section class="hero">
      <div class="hero-inner">
        <div class="hero-eyebrow">UOB AI Innovation Group · Internal Research</div>
        <h1 class="hero-title">FinDocIQ</h1>
        <p class="hero-sub">
          An agentic pipeline that reads Singapore bank CFO presentations and
          regulatory disclosures — and turns them into structured, analysis-ready Excel workbooks.
          No manual copy-paste. No lost formatting. Just clean data.
        </p>
        <div class="hero-stats">
          <div class="stat-box">
            <div class="stat-num">{total_slides}</div>
            <div class="stat-lbl">Slides extracted</div>
          </div>
          <div class="stat-box">
            <div class="stat-num">{total_pts:,}</div>
            <div class="stat-lbl">Data points</div>
          </div>
          <div class="stat-box">
            <div class="stat-num">{len(banks)}</div>
            <div class="stat-lbl">Banks covered</div>
          </div>
          <div class="stat-box">
            <div class="stat-num">${total_cost:.2f}</div>
            <div class="stat-lbl">Total API cost</div>
          </div>
        </div>
        <div class="hero-types">
          <span class="hero-types-label">Chart types handled:</span>
          {type_pills}
          <span class="hero-pill muted">+ more</span>
        </div>
      </div>
    </section>"""

def build_how_it_works():
    # ── Pillar 3 pipeline nodes ──────────────────────────────────────────────
    p3_nodes = """
      <div class="flow">

        <div class="flow-node input-node">
          <div class="fn-icon">📄</div>
          <div class="fn-title">Pillar 3 PDF</div>
          <div class="fn-sub">DBS · OCBC · UOB<br>~100 pages</div>
        </div>
        <div class="flow-arrow">↓</div>

        <div class="flow-node step-node" style="border-color:#8e44ad">
          <div class="fn-badge" style="background:#8e44ad">No API · Free</div>
          <div class="fn-step">Step 1 — Read Table of Contents</div>
          <div class="fn-title">pdfplumber extracts the TOC</div>
          <div class="fn-desc">
            Pure Python — no AI involved. Reads section titles, page numbers, and footnotes
            directly from the PDF text layer. Handles DBS Part A/B/C numbering, OCBC title overflow,
            and UOB multi-table subsections automatically.
          </div>
          <div class="fn-output">Output: section tree with start/end page for every table</div>
        </div>
        <div class="flow-arrow">↓</div>

        <div class="flow-node decision-node">
          <div class="fn-decision-icon">◆</div>
          <div class="fn-title">How many sections share this page?</div>
        </div>

        <div class="flow-branches">
          <div class="flow-branch">
            <div class="branch-label" style="background:#dbeafe;color:#1e40af">1 section, 1 page</div>
            <div class="flow-node step-node" style="border-color:#2980b9">
              <div class="fn-badge" style="background:#2980b9">Gemini 2.5 Pro · PDF + prompt</div>
              <div class="fn-step">Mode: Single</div>
              <div class="fn-title">"Extract every table on this page"</div>
              <div class="fn-desc">Sends the PDF slice (not an image) to Gemini with a structured schema. One call, one page.</div>
            </div>
          </div>
          <div class="flow-branch">
            <div class="branch-label" style="background:#fef9c3;color:#854d0e">Multiple sections, same page</div>
            <div class="flow-node step-node" style="border-color:#d97706">
              <div class="fn-badge" style="background:#d97706">Gemini 2.5 Pro · PDF + prompt</div>
              <div class="fn-step">Mode: Multiple</div>
              <div class="fn-title">"Read top-to-bottom, assign each table to its section"</div>
              <div class="fn-desc">Prompt tells Gemini to treat each section heading as a routing boundary. Tables are tagged by section ID.</div>
            </div>
          </div>
          <div class="flow-branch">
            <div class="branch-label" style="background:#fce7f3;color:#9d174d">1 section spans many pages</div>
            <div class="flow-node step-node" style="border-color:#db2777">
              <div class="fn-badge" style="background:#db2777">Gemini 2.5 Pro · PDF chunks</div>
              <div class="fn-step">Mode: Spanning</div>
              <div class="fn-title">Split into ≤2-page chunks with context carry-over</div>
              <div class="fn-desc">Each chunk gets a continuation prompt with column headers from the previous chunk so Gemini can stitch rows across page breaks without losing structure.</div>
            </div>
          </div>
        </div>

        <div class="flow-arrow">↓</div>
        <div class="flow-node decision-node">
          <div class="fn-decision-icon">◆</div>
          <div class="fn-title">Did Gemini return useful tables?</div>
          <div class="fn-sub-dec">Checks: are there tables? Do they have columns and rows? Any non-empty values?</div>
        </div>
        <div class="flow-branches flow-branches-2">
          <div class="flow-branch">
            <div class="branch-label ok-label">✓ Yes — looks good</div>
            <div class="flow-node mini-node">Proceed to render</div>
          </div>
          <div class="flow-branch">
            <div class="branch-label warn-label">✗ Empty or thin response</div>
            <div class="flow-node step-node" style="border-color:#e67e22">
              <div class="fn-badge" style="background:#e67e22">Retry · PDF + PNG image</div>
              <div class="fn-step">Image fallback</div>
              <div class="fn-title">Re-send with a rendered screenshot attached</div>
              <div class="fn-desc">PNG rendered at 2× scale alongside the original PDF. Fires once. If still thin, the result is used as-is and flagged in the audit log.</div>
            </div>
          </div>
        </div>

        <div class="flow-arrow">↓</div>
        <div class="flow-node step-node" style="border-color:#16a085">
          <div class="fn-badge" style="background:#16a085">No API · Free</div>
          <div class="fn-step">Step 3 — Render to Excel</div>
          <div class="fn-title">Write structured workbook</div>
          <div class="fn-desc">One tab per section. Hierarchy preserved (bold totals, indented sub-items). Brand colours applied. Every cell value is verbatim from the PDF — no rounding.</div>
          <div class="fn-output">Output: ocbc.xlsx / dbs.xlsx / uob.xlsx</div>
        </div>

      </div>"""

    # ── Slides pipeline nodes ────────────────────────────────────────────────
    slides_nodes = """
      <div class="flow">

        <div class="flow-node input-node">
          <div class="fn-icon">📊</div>
          <div class="fn-title">CFO Presentation PDF</div>
          <div class="fn-sub">DBS · OCBC · UOB<br>20–30 slides, charts + tables</div>
        </div>
        <div class="flow-arrow">↓</div>

        <div class="flow-node step-node" style="border-color:#8e44ad">
          <div class="fn-badge" style="background:#8e44ad">Gemini 2.5 Flash · image</div>
          <div class="fn-step">Pass 0 — Classify</div>
          <div class="fn-title">What's on this slide?</div>
          <div class="fn-desc">
            Sends the slide as a PNG image. Gemini returns a list of chart types:
            waterfall, stacked bar, donut ring, text table, KPI grid, etc.
            Costs ~$0.0001 per slide. Result is cached — free on re-runs.
          </div>
          <div class="fn-output">Output: ["donut_dual_ring", "text_table"]</div>
        </div>
        <div class="flow-arrow">↓</div>

        <div class="flow-node decision-node">
          <div class="fn-decision-icon">◆</div>
          <div class="fn-title">Does the slide contain any visual chart?</div>
          <div class="fn-sub-dec">waterfall · stacked bar · donut ring · trend line · KPI · pie</div>
        </div>

        <div class="flow-branches flow-branches-2">
          <div class="flow-branch">
            <div class="branch-label" style="background:#fff7ed;color:#c2410c">Yes — has charts</div>
            <div class="flow-node step-node" style="border-color:#e67e22">
              <div class="fn-badge" style="background:#e67e22">Gemini 2.5 Flash · image + schema</div>
              <div class="fn-step">Single-Pass extraction</div>
              <div class="fn-title">Image stays present throughout the whole call</div>
              <div class="fn-desc">
                Gemini reads the visual and fills the JSON schema in one step.
                No intermediate text description — so there's no risk of a
                correct reading being lost in transcription.
                Chart contracts (per type) guide how to read ring labels,
                bar colours, and waterfall signs.
              </div>
              <div class="fn-prompt-link" onclick="togglePrompt('sp-prompt')">
                View single-pass prompt ▾
              </div>
              <div class="fn-prompt-box" id="sp-prompt" style="display:none">
Extract ALL financial data from this bank CFO presentation slide.

Donut dual ring: trace each period label's callout line to identify
which ring it points to. Assign period from what you read on the slide.
Do NOT assume inner=earlier or outer=later.

Waterfall: every bridge bar needs sign="+" or "-". Read the colour
legend. Verify: start + sum(signed deltas) = end_value.

value field: ALWAYS verbatim as printed. "5,948" not 5948.
Return ONLY the JSON object. No markdown.</div>
            </div>
          </div>
          <div class="flow-branch">
            <div class="branch-label" style="background:#eff6ff;color:#1e40af">No — text table only</div>
            <div class="flow-node step-node" style="border-color:#2980b9">
              <div class="fn-badge" style="background:#2980b9">Gemini 2.5 Flash · image → text</div>
              <div class="fn-step">Multi-Pass extraction</div>
              <div class="fn-title">Describe first, then parse — separately</div>
              <div class="fn-desc">
                <strong>Pass 1</strong> — image sent. Gemini describes the table structure
                in plain English: column headers, row hierarchy (bold = total,
                indented = sub-item), and pre-maps every cell to schema fields.<br><br>
                <strong>Pass 2</strong> — no image. Pure text-to-JSON transcription
                of Pass 1's pre-mapping. Separating visual reading from schema
                filling improves hierarchy accuracy on dense P&amp;L tables.
              </div>
              <div class="fn-prompt-link" onclick="togglePrompt('mp-prompt')">
                View Pass 1 prompt ▾
              </div>
              <div class="fn-prompt-box" id="mp-prompt" style="display:none">
Examine this bank CFO presentation slide. It contains text tables.

STEP 1 — INVENTORY
List every printed number with its row label and column header.

STEP 2 — STRUCTURE
Column headers, row count, visual conventions:
  bold rows = totals?   indented = sub-items?
  parentheses = negative?   dash = zero?

STEP 3 — VERIFY
Do sub-items sum to their parent total? Write the check explicitly.

STEP 4 — PRE-MAP
Map every cell:
  series="Net Interest Income" period="FY25" value="5,948"
  row_type="data" level=1</div>
            </div>
          </div>
        </div>

        <div class="flow-arrow">↓</div>
        <div class="flow-node decision-node">
          <div class="fn-decision-icon">◆</div>
          <div class="fn-title">Validation — does the arithmetic check out?</div>
          <div class="fn-sub-dec">Waterfall balance · blank labels · illegible values</div>
        </div>
        <div class="flow-branches flow-branches-2">
          <div class="flow-branch">
            <div class="branch-label ok-label">✓ Passes</div>
            <div class="flow-node mini-node">Save to audit and proceed</div>
          </div>
          <div class="flow-branch">
            <div class="branch-label warn-label">✗ Errors found</div>
            <div class="flow-node step-node" style="border-color:#e67e22">
              <div class="fn-badge" style="background:#e67e22">1× correction call</div>
              <div class="fn-step">Correction pass</div>
              <div class="fn-title">Re-extract only the broken elements</div>
              <div class="fn-desc">Sends the specific validation errors back to Gemini with the original context. Only fires once. If it doesn't improve, original output is kept and errors are flagged.</div>
            </div>
          </div>
        </div>

        <div class="flow-arrow">↓</div>
        <div class="flow-node step-node" style="border-color:#16a085">
          <div class="fn-badge" style="background:#16a085">No API · Free</div>
          <div class="fn-step">Pass 3 — Render to Excel</div>
          <div class="fn-title">Wide-format pivot workbook</div>
          <div class="fn-desc">Periods become columns, series become rows. Bold totals, indented sub-items, yellow cells for chart-sourced values (verify against source). One tab per slide, one workbook per bank.</div>
          <div class="fn-output">Output: ocbc_slides.xlsx · dbs_slides.xlsx · uob_slides.xlsx</div>
        </div>

      </div>"""

    return f"""
    <section class="section" id="how">
      <div class="section-inner">
        <div class="section-eyebrow">Architecture</div>
        <h2 class="section-title">How it works</h2>
        <p class="section-sub">Two pipelines — one for regulatory disclosures, one for CFO slide decks. Toggle between them.</p>

        <div class="arch-toggle">
          <button class="arch-btn active" onclick="showArch('slides', this)">📊 CFO Slides</button>
          <button class="arch-btn" onclick="showArch('p3', this)">📁 Pillar 3 Disclosures</button>
        </div>

        <div id="arch-slides" class="arch-pane">
          <div class="arch-context">
            <strong>Input:</strong> CFO presentation PDFs (20–30 slides). Mix of charts, waterfall bridges, donut rings, and P&amp;L tables.
            Model used: <span class="model-tag">Gemini 2.5 Flash</span>
            Input format: <span class="model-tag">PNG image per slide</span>
          </div>
          {slides_nodes}
        </div>

        <div id="arch-p3" class="arch-pane" style="display:none">
          <div class="arch-context">
            <strong>Input:</strong> Regulatory PDF disclosures (~100 pages). Dense financial tables with row hierarchies, merged cells, and bank-specific layouts.
            Model used: <span class="model-tag">Gemini 2.5 Pro</span>
            Input format: <span class="model-tag">Native PDF slice (not image)</span>
            TOC step: <span class="model-tag">No API — pure Python</span>
          </div>
          {p3_nodes}
        </div>

      </div>
    </section>"""

def build_results(slides, summaries):
    # Download buttons — cost from per-slide meta, counts from slides list
    bank_stats = {}
    for s in slides:
        b = s["bank"]
        if b not in bank_stats:
            bank_stats[b] = {"n": 0, "pts": 0, "cost": 0.0}
        bank_stats[b]["n"]    += 1
        bank_stats[b]["pts"]  += s["n_pts"]
        bank_stats[b]["cost"] += s["cost"]
    # pull elapsed from summaries where available
    elapsed_map = {s.get("bank","").upper(): s.get("elapsed_human","—") for s in summaries}

    dl_btns = ""
    for bank in sorted(bank_stats):
        xl_path = OUTPUTS_DIR / f"{bank.lower()}_slides.xlsx"
        xl_b64  = b64_xlsx(xl_path)
        st      = bank_stats[bank]
        elapsed = elapsed_map.get(bank, "—")
        if xl_b64:
            dl_btns += f"""
            <div class="dl-card">
              <div class="dl-bank">{bank_dot(bank)}</div>
              <div class="dl-meta">{st['n']} slides · {st['pts']:,} data points · ${st['cost']:.4f} · {elapsed}</div>
              <a class="dl-btn" download="{bank.lower()}_slides.xlsx"
                 href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{xl_b64}">
                ↓ Download Excel
              </a>
            </div>"""

    # Slide cards — all banks
    all_types = sorted({t for s in slides for t in s["types"]})
    banks = sorted({s["bank"] for s in slides})

    filter_btns = '<button class="filter-btn active" onclick="filterSlides(\'all\',this)">All</button>'
    for b in banks:
        filter_btns += f'<button class="filter-btn" onclick="filterSlides(\'{b}\',this)">{b}</button>'
    filter_btns += '<span class="filter-sep">|</span>'
    for t in all_types:
        safe = t.replace("_","-")
        filter_btns += f'<button class="filter-btn type-filter" onclick="filterType(\'{safe}\',this)">{t.replace("_"," ")}</button>'

    cards = ""
    for s in slides:
        type_tags = " ".join(f'data-type="{t.replace("_","-")}"' for t in s["types"])
        chips_html = "".join(chip(t) for t in s["types"])
        img_html = (f'<img src="{s["image_b64"]}" class="slide-img" alt="slide {s["num"]}">'
                    if s.get("image_b64") else '<div class="slide-img-ph">no image</div>')
        branch_cls = "branch-single" if s["branch"]=="single-pass" else "branch-multi"
        err_html = (f'<span class="tag-warn">⚠ {len(s["errors"])} errors</span>'
                    if s["errors"] else '<span class="tag-ok">✓</span>')
        # mini datapoints table (first 8 rows, 3 cols)
        dp_rows = ""
        shown = {}
        for dp in s["datapoints"]:
            k = (dp.get("element_title","")[:30], dp.get("series","")[:30])
            if k in shown: continue
            shown[k] = True
            if len(shown) > 8: break
            dp_rows += f"""<tr>
              <td>{esc(dp.get("series","")[:35])}</td>
              <td>{esc(str(dp.get("period","") or ""))}</td>
              <td>{esc(dp.get("value",""))}</td>
            </tr>"""

        card_id = f"card-{s['bank']}-{s['num']}"
        cards += f"""
        <div class="slide-card" data-bank="{s['bank']}" {type_tags} id="{card_id}">
          <div class="slide-card-img">{img_html}</div>
          <div class="slide-card-body">
            <div class="slide-meta-row">
              {bank_dot(s['bank'])}
              <span class="slide-num-tag">Slide {s['num']:02d}</span>
              <span class="{branch_cls}">{s['branch']}</span>
              {err_html}
            </div>
            <div class="slide-title-text">{esc(s['title'])}</div>
            <div class="chips-row">{chips_html}</div>
            <div class="slide-stats-row">
              <span>{s['n_pts']} data points</span>
              <span>${s['cost']:.4f}</span>
            </div>
            <div class="dp-preview" id="dp-{card_id}" style="display:none">
              <table class="dp-table">
                <thead><tr><th>Series</th><th>Period</th><th>Value</th></tr></thead>
                <tbody>{dp_rows}</tbody>
              </table>
            </div>
            <button class="expand-btn" onclick="toggleDP('dp-{card_id}',this)">Show data ▾</button>
          </div>
        </div>"""

    return f"""
    <section class="section alt-bg" id="results">
      <div class="section-inner">
        <div class="section-eyebrow">Output</div>
        <h2 class="section-title">Extracted results</h2>
        <p class="section-sub">
          {len(slides)} slides across DBS, OCBC, and UOB — all extracted, validated, and ready to download.
        </p>
        <div class="dl-row">{dl_btns}</div>
        <div class="filter-bar">{filter_btns}</div>
        <div class="slides-grid" id="slides-grid">{cards}</div>
      </div>
    </section>"""

def build_experiments(exps):
    # Split pillar3 (E-01..E-06) vs slides (E-07+)
    p3   = [e for e in exps if int(e["id"][2:]) <= 6]
    sl   = [e for e in exps if int(e["id"][2:]) >= 7]

    def exp_card(e):
        out_text = esc(e["outcome"])
        learn_text = esc(e["learning"])
        out_text  = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', out_text)
        learn_text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', learn_text)
        return f"""
        <div class="exp-card">
          <div class="exp-header">
            <span class="exp-id">{e['id']}</span>
            <span class="exp-title-text">{esc(e['title'])}</span>
            {badge(e['status'])}
          </div>
          {"<div class='exp-field'><div class='exp-label'>Outcome</div><div class='exp-text'>"+out_text+"</div></div>" if e['outcome'] else ""}
          {"<div class='exp-field'><div class='exp-label'>Learning</div><div class='exp-text'>"+learn_text+"</div></div>" if e['learning'] else ""}
        </div>"""

    p3_html = "".join(exp_card(e) for e in p3)
    sl_html = "".join(exp_card(e) for e in sl)

    return f"""
    <section class="section" id="devlog">
      <div class="section-inner">
        <div class="section-eyebrow">Research Log</div>
        <h2 class="section-title">What we tried & what we learned</h2>
        <p class="section-sub">
          Every approach — successful or not — is documented. Failed experiments are as
          important as successful ones.
        </p>
        <div class="exp-tracks">
          <div class="exp-track">
            <div class="track-header">📁 Pillar 3 Pipeline</div>
            <div class="track-desc">Regulatory PDF extraction — DBS, OCBC, UOB Pillar 3, LCR, press releases</div>
            {p3_html or '<p class="muted-p">No entries found.</p>'}
          </div>
          <div class="exp-track">
            <div class="track-header">📊 CFO Slides Pipeline</div>
            <div class="track-desc">Slide deck extraction — charts, waterfalls, tables, donut rings</div>
            {sl_html or '<p class="muted-p">No entries found.</p>'}
          </div>
        </div>
      </div>
    </section>"""

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --red:    #1b4f91;
  --red2:   #1b6ec2;
  --dark:   #111827;
  --dark2:  #1f2937;
  --mid:    #374151;
  --muted:  #6b7280;
  --light:  #9ca3af;
  --border: #e5e7eb;
  --bg:     #ffffff;
  --bg2:    #f9fafb;
  --bg3:    #f3f4f6;
  --sans:   'Inter', sans-serif;
  --mono:   'IBM Plex Mono', monospace;
  --r:      10px;
}

html { scroll-behavior: smooth; }
body { font-family: var(--sans); background: var(--bg); color: var(--dark); font-size: 16px; line-height: 1.6; }

/* ── Nav ── */
.nav {
  position: sticky; top: 0; z-index: 100;
  background: rgba(255,255,255,0.95);
  backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--border);
  padding: 0 48px;
  display: flex; align-items: center; justify-content: space-between;
  height: 60px;
}
.nav-logo { font-size: 18px; font-weight: 800; color: var(--red); letter-spacing: -.02em; }
.nav-links { display: flex; gap: 32px; }
.nav-links a {
  font-size: 14px; font-weight: 500; color: var(--mid);
  text-decoration: none; transition: color .15s;
}
.nav-links a:hover { color: var(--red); }
.nav-badge {
  font-size: 11px; font-weight: 600; background: var(--red);
  color: #fff; border-radius: 20px; padding: 3px 10px;
}

/* ── Hero ── */
.hero {
  background: linear-gradient(160deg, #eff6ff 0%, #dbeafe 40%, #fff 100%);
  padding: 100px 48px 80px;
  text-align: center;
  border-bottom: 1px solid var(--border);
}
.hero-inner { max-width: 800px; margin: 0 auto; }
.hero-eyebrow {
  font-size: 13px; font-weight: 600; text-transform: uppercase;
  letter-spacing: .1em; color: var(--red); margin-bottom: 20px;
}
.hero-title {
  font-size: clamp(56px, 8vw, 88px);
  font-weight: 800; color: var(--dark);
  letter-spacing: -.04em; line-height: 1.05;
  margin-bottom: 24px;
}
.hero-sub {
  font-size: 20px; font-weight: 400; color: var(--mid);
  line-height: 1.65; max-width: 640px; margin: 0 auto 48px;
}
.hero-stats {
  display: flex; justify-content: center; gap: 0;
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--r); overflow: hidden;
  box-shadow: 0 2px 12px rgba(0,0,0,.06);
  margin-bottom: 32px;
}
.stat-box {
  flex: 1; padding: 24px 20px; text-align: center;
  border-right: 1px solid var(--border);
}
.stat-box:last-child { border-right: none; }
.stat-num { font-size: 36px; font-weight: 800; color: var(--red); letter-spacing: -.02em; }
.stat-lbl { font-size: 13px; color: var(--muted); margin-top: 4px; font-weight: 500; }
.hero-types { display: flex; align-items: center; justify-content: center; gap: 8px; flex-wrap: wrap; }
.hero-types-label { font-size: 13px; color: var(--muted); font-weight: 500; }
.hero-pill {
  background: var(--bg3); border: 1px solid var(--border);
  border-radius: 20px; padding: 4px 12px;
  font-size: 12px; font-weight: 500; color: var(--mid);
}
.hero-pill.muted { color: var(--light); }

/* ── Sections ── */
.section { padding: 96px 48px; }
.alt-bg  { background: var(--bg2); }
.section-inner { max-width: 1140px; margin: 0 auto; }
.section-eyebrow {
  font-size: 13px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .1em; color: var(--red); margin-bottom: 12px;
}
.section-title {
  font-size: clamp(32px, 4vw, 48px); font-weight: 800;
  color: var(--dark); letter-spacing: -.03em; line-height: 1.1;
  margin-bottom: 16px;
}
.section-sub {
  font-size: 18px; color: var(--muted); max-width: 580px;
  line-height: 1.65; margin-bottom: 56px;
}

/* ── Architecture toggle ── */
.arch-toggle {
  display: flex; gap: 8px; margin-bottom: 32px;
}
.arch-btn {
  padding: 10px 24px; border-radius: 8px;
  font-size: 14px; font-weight: 600; cursor: pointer;
  border: 2px solid var(--border); background: var(--bg3); color: var(--mid);
  transition: all .15s;
}
.arch-btn:hover { border-color: var(--red); color: var(--red); }
.arch-btn.active { background: var(--red); color: #fff; border-color: var(--red); }
.arch-context {
  background: var(--bg3); border: 1px solid var(--border);
  border-radius: var(--r); padding: 14px 20px;
  font-size: 13px; color: var(--mid); margin-bottom: 32px;
  display: flex; flex-wrap: wrap; gap: 12px; align-items: center;
}
.model-tag {
  background: #fff; border: 1px solid var(--border);
  border-radius: 4px; padding: 2px 8px;
  font-family: var(--mono); font-size: 12px; color: var(--dark); font-weight: 500;
}

/* ── Flow diagram ── */
.flow {
  display: flex; flex-direction: column; align-items: center; gap: 0;
  max-width: 860px; margin: 0 auto;
}
.flow-arrow {
  font-size: 22px; color: var(--light); line-height: 1; padding: 6px 0;
}
.flow-node {
  width: 100%; border-radius: 10px; padding: 20px 24px;
  border: 2px solid var(--border); background: var(--bg);
}
.input-node {
  text-align: center; background: var(--bg3); max-width: 320px;
  padding: 24px;
}
.input-node .fn-icon { font-size: 28px; margin-bottom: 8px; }
.input-node .fn-title { font-size: 16px; font-weight: 700; color: var(--dark); }
.input-node .fn-sub   { font-size: 13px; color: var(--muted); margin-top: 4px; }
.step-node  { background: var(--bg); }
.mini-node  {
  max-width: 200px; text-align: center;
  background: #f0fdf4; border: 1.5px solid #86efac;
  border-radius: 8px; padding: 10px 16px;
  font-size: 13px; font-weight: 500; color: #166534;
}
.decision-node {
  text-align: center; background: #fefce8;
  border: 2px dashed #fbbf24; max-width: 480px;
  padding: 16px 24px;
}
.fn-decision-icon { font-size: 18px; margin-bottom: 6px; }
.fn-sub-dec { font-size: 12px; color: var(--muted); margin-top: 4px; }
.fn-badge {
  display: inline-block; padding: 3px 10px; border-radius: 20px;
  font-size: 11px; font-family: var(--mono); color: #fff; font-weight: 600;
  margin-bottom: 8px;
}
.fn-step  { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: var(--muted); margin-bottom: 4px; }
.fn-title { font-size: 16px; font-weight: 700; color: var(--dark); margin-bottom: 8px; line-height: 1.3; }
.fn-desc  { font-size: 14px; color: var(--mid); line-height: 1.7; }
.fn-output {
  margin-top: 10px; font-size: 12px; font-family: var(--mono);
  color: var(--red); background: var(--bg3);
  border-radius: 4px; padding: 4px 10px; display: inline-block;
}

/* ── Prompt reveal ── */
.fn-prompt-link {
  margin-top: 12px; font-size: 12px; color: var(--red);
  cursor: pointer; font-weight: 600; display: inline-block;
}
.fn-prompt-link:hover { text-decoration: underline; }
.fn-prompt-box {
  margin-top: 8px; background: #0f172a; color: #e2e8f0;
  border-radius: 6px; padding: 14px 16px;
  font-family: var(--mono); font-size: 12px; line-height: 1.7;
  white-space: pre-wrap; word-break: break-word;
}

/* ── Branch layout ── */
.flow-branches {
  display: grid; grid-template-columns: repeat(3, 1fr);
  gap: 16px; width: 100%;
}
.flow-branches-2 {
  display: grid; grid-template-columns: 1fr 1fr;
  gap: 16px; width: 100%;
}
.flow-branch { display: flex; flex-direction: column; align-items: center; gap: 8px; }
.branch-label {
  font-size: 12px; font-weight: 600; padding: 4px 14px;
  border-radius: 20px; text-align: center;
}
.ok-label   { background: #f0fdf4; color: #166534; }
.warn-label { background: #fff7ed; color: #c2410c; }

/* ── Principles (kept for fallback) ── */
.principles-grid {
  display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
  gap: 20px;
}
.principle-card {
  background: #eff6ff; border: 1px solid #bfdbfe;
  border-radius: var(--r); padding: 24px;
}
.principle-title { font-size: 15px; font-weight: 700; color: var(--dark); margin-bottom: 8px; }
.principle-desc  { font-size: 14px; color: var(--mid); line-height: 1.65; }

.principles-grid {
  display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
  gap: 20px;
}
.principle-card {
  background: #eff6ff; border: 1px solid #bfdbfe;
  border-radius: var(--r); padding: 24px;
}
.principle-title { font-size: 15px; font-weight: 700; color: var(--dark); margin-bottom: 8px; }
.principle-desc  { font-size: 14px; color: var(--mid); line-height: 1.65; }

/* ── Downloads ── */
.dl-row { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 40px; }
.dl-card {
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--r); padding: 20px 24px;
  display: flex; align-items: center; gap: 20px;
  box-shadow: 0 1px 4px rgba(0,0,0,.05);
}
.dl-bank { }
.dl-meta { font-size: 13px; color: var(--muted); }
.dl-btn {
  background: var(--red); color: #fff;
  border-radius: 6px; padding: 9px 18px;
  font-size: 13px; font-weight: 600;
  text-decoration: none; white-space: nowrap;
  transition: background .15s;
}
.dl-btn:hover { background: var(--red2); }

/* ── Filter bar ── */
.filter-bar {
  display: flex; align-items: center; gap: 8px;
  flex-wrap: wrap; margin-bottom: 32px;
}
.filter-btn {
  background: var(--bg3); border: 1px solid var(--border);
  border-radius: 20px; padding: 5px 14px;
  font-size: 13px; font-weight: 500; color: var(--mid);
  cursor: pointer; transition: all .15s;
}
.filter-btn:hover  { border-color: var(--red); color: var(--red); }
.filter-btn.active { background: var(--red); color: #fff; border-color: var(--red); }
.filter-sep { color: var(--border); font-size: 18px; margin: 0 4px; }

/* ── Slide cards ── */
.slides-grid {
  display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
  gap: 20px;
}
.slide-card {
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--r); overflow: hidden;
  transition: box-shadow .2s, transform .2s;
}
.slide-card:hover { box-shadow: 0 8px 24px rgba(0,0,0,.1); transform: translateY(-2px); }
.slide-card.hidden { display: none; }
.slide-card-img .slide-img {
  width: 100%; display: block; background: #f8f8f8;
}
.slide-img-ph {
  width: 100%; aspect-ratio: 16/9; background: var(--bg3);
  display: flex; align-items: center; justify-content: center;
  color: var(--light); font-size: 12px;
}
.slide-card-body { padding: 16px; }
.slide-meta-row  { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 8px; }
.slide-title-text { font-size: 14px; font-weight: 600; color: var(--dark); line-height: 1.4; margin-bottom: 10px; }
.chips-row       { display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 10px; }
.slide-stats-row { display: flex; gap: 16px; font-size: 12px; color: var(--muted); margin-bottom: 12px; }

.bank-dot {
  display: inline-block; padding: 2px 8px; border-radius: 4px;
  font-size: 11px; font-family: var(--mono); color: #fff; font-weight: 600;
}
.slide-num-tag {
  font-family: var(--mono); font-size: 11px; color: var(--muted);
  background: var(--bg3); border-radius: 4px; padding: 1px 6px;
}
.branch-single {
  font-size: 11px; font-weight: 600; color: #d97706;
  background: #fef3c7; border-radius: 4px; padding: 1px 6px;
}
.branch-multi {
  font-size: 11px; font-weight: 600; color: #1d4ed8;
  background: #eff6ff; border-radius: 4px; padding: 1px 6px;
}
.tag-ok   { font-size: 11px; color: #16a34a; font-weight: 600; }
.tag-warn { font-size: 11px; color: #d97706; font-weight: 600; }

.chip {
  display: inline-block; padding: 2px 7px; border-radius: 4px;
  font-size: 10px; font-family: var(--mono); color: #fff;
}

/* ── Data preview table ── */
.expand-btn {
  background: none; border: 1px solid var(--border);
  border-radius: 6px; padding: 5px 12px;
  font-size: 12px; color: var(--muted); cursor: pointer;
  width: 100%; text-align: left; transition: all .15s;
}
.expand-btn:hover { border-color: var(--red); color: var(--red); }
.dp-preview { margin-bottom: 10px; overflow-x: auto; }
.dp-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.dp-table th {
  background: var(--bg3); color: var(--muted);
  font-size: 10px; text-transform: uppercase; letter-spacing: .05em;
  padding: 6px 8px; text-align: left; border-bottom: 1px solid var(--border);
}
.dp-table td {
  padding: 5px 8px; border-bottom: 1px solid var(--bg3);
  color: var(--dark2); max-width: 160px;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}

/* ── Experiment cards ── */
.exp-tracks { display: grid; grid-template-columns: 1fr 1fr; gap: 40px; }
@media(max-width:900px){ .exp-tracks { grid-template-columns: 1fr; } }
.exp-track {}
.track-header {
  font-size: 18px; font-weight: 700; color: var(--dark);
  margin-bottom: 6px;
}
.track-desc { font-size: 14px; color: var(--muted); margin-bottom: 24px; line-height: 1.6; }
.exp-card {
  background: #fff; border: 1px solid var(--border);
  border-radius: var(--r); padding: 20px 22px;
  margin-bottom: 14px;
}
.exp-header { display: flex; align-items: center; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; }
.exp-id    { font-family: var(--mono); font-size: 12px; font-weight: 600; color: var(--red); }
.exp-title-text { font-size: 14px; font-weight: 600; color: var(--dark); flex: 1; }
.badge {
  display: inline-block; padding: 2px 8px; border-radius: 4px;
  font-size: 10px; font-family: var(--mono); color: #fff; font-weight: 700;
}
.exp-field { margin-bottom: 12px; }
.exp-field:last-child { margin-bottom: 0; }
.exp-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: var(--muted); margin-bottom: 4px; }
.exp-text  { font-size: 13px; color: var(--mid); line-height: 1.7; }

/* ── Footer ── */
.footer {
  border-top: 1px solid var(--border);
  padding: 40px 48px;
  display: flex; justify-content: space-between; align-items: center;
  font-size: 13px; color: var(--muted);
}
.footer-logo { font-weight: 800; color: var(--red); font-size: 15px; }
.muted-p { font-size: 14px; color: var(--muted); font-style: italic; }
"""

# ---------------------------------------------------------------------------
# JS
# ---------------------------------------------------------------------------

JS = """
// ── Architecture toggle ──
function showArch(which, btn) {
  document.querySelectorAll('.arch-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.arch-pane').forEach(p => p.style.display = 'none');
  btn.classList.add('active');
  document.getElementById('arch-' + which).style.display = 'block';
}

// ── Prompt reveal ──
function togglePrompt(id) {
  const el = document.getElementById(id);
  const link = el.previousElementSibling;
  if (el.style.display === 'none') {
    el.style.display = 'block';
    link.textContent = link.textContent.replace('▾','▴');
  } else {
    el.style.display = 'none';
    link.textContent = link.textContent.replace('▴','▾');
  }
}

let activeBank = 'all';
let activeType = null;

function filterSlides(bank, btn) {
  activeBank = bank;
  activeType = null;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  applyFilter();
}

function filterType(type, btn) {
  if (activeType === type) {
    activeType = null;
    btn.classList.remove('active');
  } else {
    activeType = type;
    document.querySelectorAll('.type-filter').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }
  applyFilter();
}

function applyFilter() {
  document.querySelectorAll('.slide-card').forEach(card => {
    const bankMatch = activeBank === 'all' || card.dataset.bank === activeBank;
    const typeMatch = !activeType || card.hasAttribute('data-type-' + activeType) || card.dataset['type'] === activeType;
    // Check all data-type-* attributes
    let hasType = !activeType;
    if (activeType) {
      for (const attr of card.attributes) {
        if (attr.name.startsWith('data-type') && attr.value === activeType) {
          hasType = true; break;
        }
      }
    }
    card.classList.toggle('hidden', !(bankMatch && hasType));
  });
}

function toggleDP(id, btn) {
  const el = document.getElementById(id);
  if (el.style.display === 'none') {
    el.style.display = 'block';
    btn.textContent = 'Hide data ▴';
  } else {
    el.style.display = 'none';
    btn.textContent = 'Show data ▾';
  }
}

// Smooth scroll for nav links
document.querySelectorAll('a[href^="#"]').forEach(a => {
  a.addEventListener('click', e => {
    e.preventDefault();
    document.querySelector(a.getAttribute('href'))?.scrollIntoView({behavior:'smooth'});
  });
});
"""

# ---------------------------------------------------------------------------
# BUILD
# ---------------------------------------------------------------------------

def build() -> str:
    slides    = load_all_slides()
    summaries = load_run_summaries()
    contracts = load_contracts()
    exps      = load_devlog_experiments()
    gen_ts    = datetime.now().strftime("%Y-%m-%d %H:%M")

    hero   = build_hero(slides, summaries)
    how    = build_how_it_works()
    res    = build_results(slides, summaries)
    devlog = build_experiments(exps)

    total = len(slides)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>FinDocIQ — Financial Document Intelligence</title>
  <style>{CSS}</style>
</head>
<body>

<nav class="nav">
  <div class="nav-logo">FinDocIQ</div>
  <div class="nav-links">
    <a href="#how">How it works</a>
    <a href="#results">Results</a>
    <a href="#devlog">Research log</a>
  </div>
  <span class="nav-badge">Internal Demo</span>
</nav>

{hero}
{how}
{res}
{devlog}

<footer class="footer">
  <div class="footer-logo">FinDocIQ</div>
  <div>UOB AI Innovation Group · Generated {gen_ts} · {total} slides extracted</div>
</footer>

<script>{JS}</script>
</body>
</html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--open", action="store_true")
    ap.add_argument("--out",  default=None)
    args = ap.parse_args()

    html     = build()
    out_path = args.out or str(DEMO_DIR / "index.html")
    Path(out_path).write_text(html, encoding="utf-8")
    size_kb = Path(out_path).stat().st_size / 1024
    print(f"✓ Demo site → {out_path}  ({size_kb:.0f} KB)")
    if args.open:
        webbrowser.open(f"file://{Path(out_path).resolve()}")

if __name__ == "__main__":
    main()
