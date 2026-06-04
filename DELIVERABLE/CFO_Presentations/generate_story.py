"""
generate_story.py — FinDocIQ living project story report.

Usage:
    python generate_story.py              # auto-discovers everything
    python generate_story.py --open       # open in browser after generating
    python generate_story.py --out PATH   # custom output path
"""
from __future__ import annotations
import re, json, base64, argparse, webbrowser
from datetime import datetime
from pathlib import Path


SCRIPT_DIR  = Path(__file__).parent.resolve()
DELIVERABLE = SCRIPT_DIR.parent
MDS_DIR     = DELIVERABLE.parent / "MDs"
OUTPUTS_DIR = DELIVERABLE / "outputs" / "CFO_Presentation"
AUDIT_DIR   = OUTPUTS_DIR / "audit"
REPORTS_DIR = DELIVERABLE / "demo"

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def read_file(path) -> str:
    p = Path(path)
    return p.read_text(encoding="utf-8") if p.exists() else ""

def b64_img(path) -> str | None:
    p = Path(path)
    if not p.exists():
        return None
    data = base64.b64encode(p.read_bytes()).decode()
    ext  = p.suffix.lower().lstrip(".")
    mime = {"png": "image/png", "jpg": "image/jpeg"}.get(ext, "image/png")
    return f"data:{mime};base64,{data}"

def b64_file(path) -> str | None:
    p = Path(path)
    if not p.exists():
        return None
    return base64.b64encode(p.read_bytes()).decode()

def load_json(path):
    p = Path(path)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text())
    except Exception:
        return None

def esc(s: str) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# ---------------------------------------------------------------------------
# DATA LOADERS
# ---------------------------------------------------------------------------

def find_audit_dirs() -> dict[str, Path]:
    result = {}
    if not AUDIT_DIR.exists():
        return result
    for d in sorted(AUDIT_DIR.iterdir()):
        if d.is_dir():
            slug = d.name.split("_")[0].upper()
            result[slug] = d
    return result

def load_slide(audit_dir: Path, slide_num: int) -> dict | None:
    sd = audit_dir / f"slide_{slide_num:02d}"
    if not sd.exists():
        return None

    meta   = load_json(sd / "meta.json") or {}
    dps    = load_json(sd / "datapoints.json") or []
    types  = load_json(sd / "element_types.json") or []

    # Skip "none" slides
    if types == ["none"] or not types:
        return None

    slide_title = ""
    if dps:
        slide_title = dps[0].get("slide_title", "")

    img_b64 = b64_img(sd / f"slide_{slide_num:02d}.png") or b64_img(sd / "slide.png")

    # Determine branch used
    usages = meta.get("usages", [])
    passes = [str(u.get("pass","")) for u in usages]
    if "1s" in passes:
        branch = "single-pass"
    elif passes:
        branch = "multi-pass"
    else:
        branch = "resumed"

    return {
        "num":        slide_num,
        "title":      slide_title or f"Slide {slide_num}",
        "types":      types,
        "meta":       meta,
        "n_pts":      meta.get("n_points", len(dps)),
        "cost":       meta.get("cost_usd", 0.0),
        "errors":     meta.get("validation_errors") or [],
        "branch":     branch,
        "image_b64":  img_b64,
        "datapoints": dps,
    }

def load_run_summaries() -> list[dict]:
    if not OUTPUTS_DIR.exists():
        return []
    summaries = []
    for f in OUTPUTS_DIR.glob("*_run_summary.json"):
        s = load_json(f)
        if s:
            summaries.append(s)
    return sorted(summaries, key=lambda x: x.get("generated_at", ""))

def load_excel_b64(bank: str) -> str | None:
    slug = bank.lower()
    p = OUTPUTS_DIR / f"{slug}_slides.xlsx"
    return b64_file(p)

# ---------------------------------------------------------------------------
# DEVLOG PARSER
# ---------------------------------------------------------------------------

def parse_devlog(md: str) -> dict:
    pillar3, slides = [], []

    exp_re = re.compile(
        r"###\s+(E-\d+)\s+[—–-]+\s+(.+?)\s+[●•]\s+(\w[\w\s/]+?)(?:\s*/[^\n]*)?\n",
        re.IGNORECASE,
    )

    def get_field(body: str, field: str) -> str:
        m = re.search(rf"\|\s*\*\*{field}\*\*\s*\|\s*(.+?)(?=\n\||\Z)", body, re.DOTALL)
        return m.group(1).strip()[:500] if m else ""

    matches = list(exp_re.finditer(md))
    slides_boundary = md.find("## 7.")

    for i, m in enumerate(matches):
        exp_id = m.group(1)
        title  = m.group(2).strip()
        status_raw = m.group(3).strip().upper()

        if   "ADOPT"    in status_raw: status = "adopted"
        elif "PARTIAL"  in status_raw: status = "partial"
        elif "INSIGHT"  in status_raw: status = "insight"
        elif "IMPLEMENT" in status_raw: status = "implemented"
        elif "DESIGN"   in status_raw: status = "insight"
        else:                          status = "failed"

        body_start = m.end()
        body_end   = matches[i+1].start() if i+1 < len(matches) else len(md)
        body       = md[body_start:body_end].strip()

        exp = {
            "id": exp_id, "title": title, "status": status,
            "hypothesis": get_field(body, "Hypothesis"),
            "outcome":    get_field(body, "Outcome"),
            "learning":   get_field(body, "Learning"),
            "decision":   get_field(body, "Decision"),
        }

        num = int(exp_id[2:])
        is_slide_exp = num >= 7 or (slides_boundary > 0 and m.start() > slides_boundary)
        (slides if is_slide_exp else pillar3).append(exp)

    return {"pillar3": pillar3, "slides": slides}

# ---------------------------------------------------------------------------
# HTML FRAGMENTS
# ---------------------------------------------------------------------------

TYPE_COLOURS = {
    "waterfall":                "#e74c3c",
    "stacked_bar":              "#3498db",
    "stacked_bar_with_overlay": "#2980b9",
    "trend_line":               "#27ae60",
    "kpi_grid":                 "#8e44ad",
    "text_table":               "#2c3e50",
    "pie":                      "#f39c12",
    "donut_dual_ring":          "#e67e22",
    "npa_movement_table":       "#16a085",
    "npa_movement_table":       "#16a085",
}

def type_chip(t: str) -> str:
    bg = TYPE_COLOURS.get(t, "#5a6a7a")
    label = t.replace("_", " ")
    return f'<span class="chip" style="background:{bg}">{label}</span>'

def badge(status: str) -> str:
    colours = {
        "failed":      "#e74c3c", "adopted":     "#27ae60",
        "partial":     "#f39c12", "insight":     "#8e44ad",
        "implemented": "#2980b9",
    }
    labels = {
        "failed": "FAILED", "adopted": "ADOPTED", "partial": "PARTIAL",
        "insight": "INSIGHT", "implemented": "IMPLEMENTED",
    }
    bg = colours.get(status, "#7f8c8d")
    lb = labels.get(status, status.upper())
    return f'<span class="badge" style="background:{bg}">{lb}</span>'

def render_run_panel(summaries: list[dict]) -> str:
    if not summaries:
        return '<p class="muted">No run summaries found.</p>'

    rows = ""
    for s in summaries:
        bank    = s.get("bank", "—")
        doc     = s.get("document", "—")
        slides  = s.get("slides_processed", "—")
        calls   = s.get("api_calls", "—")
        inp     = s.get("input_tokens", 0)
        out     = s.get("output_tokens", 0)
        cost    = s.get("est_cost_usd", 0)
        elapsed = s.get("elapsed_human", "—")
        date    = (s.get("generated_at") or "")[:16].replace("T", " ")

        dl_b64  = load_excel_b64(bank)
        dl_btn  = ""
        if dl_b64:
            fname = Path(s.get("output_file", f"{bank.lower()}_slides.xlsx")).name
            dl_btn = f'<a class="dl-btn" download="{fname}" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{dl_b64}">⬇ Download Excel</a>'

        rows += f"""
        <div class="run-card">
          <div class="run-header">
            <span class="bank-tag bank-{bank.lower()}">{bank}</span>
            <span class="run-doc">{esc(doc)}</span>
            <span class="run-date">{date}</span>
          </div>
          <div class="run-stats">
            <div class="stat"><div class="stat-val">{slides}</div><div class="stat-lbl">Slides</div></div>
            <div class="stat"><div class="stat-val">{calls}</div><div class="stat-lbl">API Calls</div></div>
            <div class="stat"><div class="stat-val">{inp:,}</div><div class="stat-lbl">Input tokens</div></div>
            <div class="stat"><div class="stat-val">{out:,}</div><div class="stat-lbl">Output tokens</div></div>
            <div class="stat"><div class="stat-val cost">${cost:.4f}</div><div class="stat-lbl">Est. Cost</div></div>
            <div class="stat"><div class="stat-val">{elapsed}</div><div class="stat-lbl">Elapsed</div></div>
          </div>
          {dl_btn}
        </div>"""

    return rows

def render_slide_grid(audit_dirs: dict[str, Path]) -> str:
    html = ""
    for bank, adir in audit_dirs.items():
        slides = []
        for sd in sorted(adir.iterdir()):
            m = re.match(r"slide_(\d+)", sd.name)
            if not m:
                continue
            s = load_slide(adir, int(m.group(1)))
            if s:
                slides.append(s)

        if not slides:
            continue

        cards = ""
        for s in slides:
            chips   = " ".join(type_chip(t) for t in s["types"])
            err_tag = (f'<span class="tag warn">⚠ {len(s["errors"])} errors</span>'
                       if s["errors"] else '<span class="tag ok">✓ passed</span>')
            branch_tag = (f'<span class="tag visual">single-pass</span>'
                          if s["branch"] == "single-pass"
                          else f'<span class="tag text">multi-pass</span>')

            img_html = (f'<img src="{s["image_b64"]}" class="slide-thumb" alt="Slide {s["num"]}">'
                        if s.get("image_b64")
                        else '<div class="slide-thumb-missing">no image</div>')

            cards += f"""
            <div class="slide-card">
              <div class="slide-num">Slide {s["num"]:02d}</div>
              <div class="slide-title">{esc(s["title"])}</div>
              {img_html}
              <div class="slide-meta">
                <div class="slide-chips">{chips}</div>
                <div class="slide-tags">
                  {branch_tag}
                  <span class="tag mono">{s["n_pts"]} pts</span>
                  <span class="tag mono cost">${s["cost"]:.4f}</span>
                  {err_tag}
                </div>
              </div>
            </div>"""

        html += f"""
        <div class="section-block">
          <div class="section-title">
            <span class="bank-tag bank-{bank.lower()}">{bank}</span>
            &nbsp;Extracted Slides — {len(slides)} slides
          </div>
          <div class="slide-grid">{cards}</div>
        </div>"""

    return html or '<p class="muted">No audit data found.</p>'

def render_devlog(devlog: dict) -> str:
    def exp_card(exp: dict) -> str:
        fields = []
        for label, key in [("Hypothesis","hypothesis"),("Outcome","outcome"),
                           ("Decision","decision"),("Learning","learning")]:
            val = exp.get(key, "")
            if not val:
                continue
            # Break long text at sentence boundaries for readability
            text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', esc(val))
            # Convert (1)/(2) numbered items and bullet lists into actual list items
            text = re.sub(r'\((\d+)\)\s+', r'<br><strong>(\1)</strong> ', text)
            fields.append(f"""
            <div class="exp-field">
              <div class="exp-field-lbl">{label}</div>
              <div class="exp-field-val">{text}</div>
            </div>""")

        return f"""
        <div class="exp-card">
          <div class="exp-header">
            <span class="exp-id">{exp["id"]}</span>
            <span class="exp-title">{esc(exp["title"])}</span>
            {badge(exp["status"])}
          </div>
          {"".join(fields)}
        </div>"""

    p3 = "".join(exp_card(e) for e in devlog["pillar3"])
    sl = "".join(exp_card(e) for e in devlog["slides"])

    return f"""
    <div class="devlog-group">
      <div class="group-header p3-header">
        <span>📁 Pillar 3 Pipeline</span>
        <span class="group-count">{len(devlog["pillar3"])} experiments</span>
      </div>
      <div class="group-desc">
        Regulatory PDF extraction (DBS / OCBC / UOB Pillar 3, LCR, press releases).
        Pass 1 = deterministic TOC extraction. Pass 2 = Gemini Vision table extraction.
      </div>
      {p3 or '<p class="muted">No entries found.</p>'}
    </div>

    <div class="devlog-group" style="margin-top:48px">
      <div class="group-header sl-header">
        <span>📊 CFO Slides Pipeline</span>
        <span class="group-count">{len(devlog["slides"])} experiments</span>
      </div>
      <div class="group-desc">
        CFO presentation slide extraction. Hybrid architecture: single-pass for visual
        charts, multi-pass for text tables.
      </div>
      {sl or '<p class="muted">No entries found.</p>'}
    </div>"""

def render_pipeline_diagram() -> str:
    return """
    <div class="pipeline">
      <div class="pipe-row">

        <div class="pipe-box input-box">
          <div class="pipe-icon">📄</div>
          <div class="pipe-name">CFO Slide PDF</div>
          <div class="pipe-desc">DBS · OCBC · UOB</div>
        </div>

        <div class="pipe-arrow">→</div>

        <div class="pipe-box pass0-box">
          <div class="pipe-pass">Pass 0</div>
          <div class="pipe-name">Classify</div>
          <div class="pipe-desc">Identify chart types<br>~150 tokens</div>
        </div>

        <div class="pipe-arrow">→</div>

        <div class="pipe-split">
          <div class="pipe-branch">
            <div class="branch-label visual-label">Visual slides</div>
            <div class="pipe-box visual-box">
              <div class="pipe-pass">Single Pass</div>
              <div class="pipe-name">Image + Schema</div>
              <div class="pipe-desc">One call. Image stays<br>present throughout.</div>
            </div>
          </div>
          <div class="pipe-branch-sep">or</div>
          <div class="pipe-branch">
            <div class="branch-label text-label">Text tables</div>
            <div class="pipe-box text-box">
              <div class="pipe-pass">Pass 1</div>
              <div class="pipe-name">Describe</div>
              <div class="pipe-desc">Image → plain text.<br>Pre-map to schema fields.</div>
            </div>
            <div class="pipe-down-arrow">↓</div>
            <div class="pipe-box text-box">
              <div class="pipe-pass">Pass 2</div>
              <div class="pipe-name">Parse</div>
              <div class="pipe-desc">Text only → JSON.<br>No image. Pure transcription.</div>
            </div>
          </div>
        </div>

        <div class="pipe-arrow">→</div>

        <div class="pipe-box val-box">
          <div class="pipe-pass">Validate</div>
          <div class="pipe-name">Arithmetic Check</div>
          <div class="pipe-desc">Waterfall balance.<br>Self-check strings.</div>
        </div>

        <div class="pipe-arrow">→</div>

        <div class="pipe-box render-box">
          <div class="pipe-pass">Pass 3</div>
          <div class="pipe-name">Render</div>
          <div class="pipe-desc">Wide-format Excel.<br>Per-slide tab.</div>
        </div>

        <div class="pipe-arrow">→</div>

        <div class="pipe-box output-box">
          <div class="pipe-icon">📊</div>
          <div class="pipe-name">Excel Workbook</div>
          <div class="pipe-desc">Wide format · Verified</div>
        </div>

      </div>

      <div class="pipe-principles">
        <div class="principle">👁 Visual slides: image present throughout — no schema pressure during reading</div>
        <div class="principle">📋 Text tables: structured decomposition catches hierarchy and arithmetic errors</div>
        <div class="principle">📜 Chart contracts teach Gemini the data model before reading values</div>
        <div class="principle">💾 Full audit trail per slide — resume is free, no re-billing</div>
      </div>
    </div>"""

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------

CSS = """
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg:     #0c1016;
  --bg2:    #111820;
  --bg3:    #18222e;
  --bg4:    #1e2b3a;
  --border: #243040;
  --text:   #dde4ef;
  --text2:  #8498b0;
  --text3:  #3d5068;
  --accent: #3d8bff;
  --green:  #3ecf72;
  --red:    #f05252;
  --orange: #f08030;
  --purple: #9b6dff;
  --mono:   'IBM Plex Mono', monospace;
  --sans:   'IBM Plex Sans', sans-serif;
  --r:      8px;
}

html { scroll-behavior: smooth; }
body {
  font-family: var(--sans);
  background: var(--bg);
  color: var(--text);
  font-size: 14px;
  line-height: 1.7;
}

/* ─── Header ─── */
.site-header {
  background: linear-gradient(135deg, #0d1e35, #060d14);
  border-bottom: 1px solid var(--border);
  padding: 32px 48px 0;
}
.header-top {
  display: flex;
  align-items: flex-end;
  justify-content: space-between;
  margin-bottom: 24px;
}
.logo { font-size: 24px; font-weight: 600; color: var(--accent); letter-spacing: -.02em; }
.logo-sub { font-size: 13px; color: var(--text2); margin-top: 4px; }
.header-meta { font-family: var(--mono); font-size: 11px; color: var(--text3); text-align: right; }

/* ─── Tabs ─── */
.tab-bar { display: flex; }
.tab {
  padding: 12px 24px;
  font-size: 13px; font-weight: 500;
  color: var(--text2);
  cursor: pointer;
  border-bottom: 2px solid transparent;
  margin-bottom: -1px;
  transition: color .15s, border-color .15s;
  user-select: none;
}
.tab:hover { color: var(--text); }
.tab.active { color: var(--accent); border-bottom-color: var(--accent); }

.tab-content { display: none; padding: 40px 48px; max-width: 1180px; }
.tab-content.active { display: block; }

/* ─── Sections ─── */
.section-block { margin-bottom: 56px; }
.section-title {
  font-size: 11px; font-weight: 600;
  text-transform: uppercase; letter-spacing: .1em;
  color: var(--text3);
  margin-bottom: 20px;
  padding-bottom: 10px;
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 10px;
}

/* ─── Bank tag ─── */
.bank-tag {
  display: inline-block; padding: 2px 9px; border-radius: 4px;
  font-size: 11px; font-family: var(--mono); color: #fff; font-weight: 500;
}
.bank-ocbc { background: #cc0000; }
.bank-uob  { background: #1b6ec2; }
.bank-dbs  { background: #cc0000; }

/* ─── Run cards ─── */
.run-card {
  background: var(--bg2);
  border: 1px solid var(--border);
  border-radius: var(--r);
  padding: 24px 28px;
  margin-bottom: 20px;
}
.run-header {
  display: flex; align-items: center; gap: 12px;
  margin-bottom: 20px; flex-wrap: wrap;
}
.run-doc { font-size: 15px; font-weight: 500; color: var(--text); flex: 1; }
.run-date { font-family: var(--mono); font-size: 11px; color: var(--text3); }
.run-stats {
  display: flex; gap: 0;
  background: var(--bg3); border: 1px solid var(--border);
  border-radius: var(--r); overflow: hidden;
  margin-bottom: 20px;
}
.stat {
  flex: 1; padding: 14px 16px; text-align: center;
  border-right: 1px solid var(--border);
}
.stat:last-child { border-right: none; }
.stat-val {
  font-family: var(--mono); font-size: 18px; font-weight: 500;
  color: var(--text); margin-bottom: 4px;
}
.stat-val.cost { color: #64a8ff; }
.stat-lbl { font-size: 10px; text-transform: uppercase; letter-spacing: .06em; color: var(--text3); }
.dl-btn {
  display: inline-block;
  padding: 9px 20px;
  background: var(--accent); color: #fff;
  border-radius: 6px; font-size: 13px; font-weight: 500;
  text-decoration: none;
  transition: background .15s;
}
.dl-btn:hover { background: #5aa0ff; }

/* ─── Slide grid ─── */
.slide-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 20px;
}
.slide-card {
  background: var(--bg2);
  border: 1px solid var(--border);
  border-radius: var(--r);
  padding: 16px;
  display: flex; flex-direction: column; gap: 10px;
}
.slide-num {
  font-family: var(--mono); font-size: 11px; color: var(--text3);
}
.slide-title {
  font-size: 13px; font-weight: 500; color: var(--text); line-height: 1.4;
}
.slide-thumb {
  width: 100%; border-radius: 4px;
  border: 1px solid var(--border); background: #fff;
}
.slide-thumb-missing {
  width: 100%; aspect-ratio: 16/9; background: var(--bg3);
  border: 1px dashed var(--border); border-radius: 4px;
  display: flex; align-items: center; justify-content: center;
  color: var(--text3); font-size: 11px;
}
.slide-meta { display: flex; flex-direction: column; gap: 8px; }
.slide-chips { display: flex; flex-wrap: wrap; gap: 4px; }
.slide-tags  { display: flex; flex-wrap: wrap; gap: 4px; }

/* ─── Chips & tags ─── */
.chip {
  display: inline-block; padding: 2px 7px; border-radius: 4px;
  font-size: 10px; font-family: var(--mono); color: #fff;
}
.tag {
  display: inline-block; padding: 2px 8px; border-radius: 20px;
  font-size: 10px; font-family: var(--mono);
  background: var(--bg3); border: 1px solid var(--border); color: var(--text2);
}
.tag.mono   { color: var(--text2); }
.tag.cost   { color: #64a8ff; }
.tag.ok     { color: var(--green); border-color: rgba(62,207,114,.3); }
.tag.warn   { color: var(--orange); border-color: rgba(240,128,48,.3); }
.tag.visual { color: var(--orange); border-color: rgba(240,128,48,.3); background: rgba(240,128,48,.08); }
.tag.text   { color: var(--accent); border-color: rgba(61,139,255,.3); background: rgba(61,139,255,.08); }

/* ─── Pipeline diagram ─── */
.pipeline {
  background: var(--bg2); border: 1px solid var(--border);
  border-radius: var(--r); padding: 32px;
}
.pipe-row {
  display: flex; align-items: center; gap: 10px;
  flex-wrap: wrap; justify-content: center;
  margin-bottom: 28px;
}
.pipe-box {
  text-align: center; padding: 14px 16px;
  border: 1px solid var(--border); border-radius: var(--r);
  min-width: 110px;
}
.pipe-icon  { font-size: 20px; margin-bottom: 6px; }
.pipe-pass  { font-family: var(--mono); font-size: 10px; color: var(--text3); margin-bottom: 4px; }
.pipe-name  { font-size: 12px; font-weight: 600; color: var(--text); margin-bottom: 4px; }
.pipe-desc  { font-size: 10px; color: var(--text2); line-height: 1.5; }
.input-box, .output-box { background: var(--bg3); }
.pass0-box  { background: rgba(155,109,255,.08); border-color: rgba(155,109,255,.25); }
.visual-box { background: rgba(240,128,48,.08);  border-color: rgba(240,128,48,.25); }
.text-box   { background: rgba(61,139,255,.08);  border-color: rgba(61,139,255,.25); }
.val-box    { background: rgba(62,207,114,.08);  border-color: rgba(62,207,114,.25); }
.render-box { background: rgba(240,192,64,.08);  border-color: rgba(240,192,64,.25); }
.pipe-arrow      { color: var(--text3); font-size: 18px; flex-shrink: 0; }
.pipe-down-arrow { color: var(--text3); font-size: 14px; text-align: center; margin: 4px 0; }
.pipe-split {
  display: flex; align-items: center; gap: 16px;
}
.pipe-branch        { display: flex; flex-direction: column; align-items: center; gap: 4px; }
.pipe-branch-sep    { font-size: 11px; color: var(--text3); align-self: center; }
.branch-label       {
  font-size: 10px; font-weight: 600; text-transform: uppercase;
  letter-spacing: .06em; padding: 2px 10px; border-radius: 10px; margin-bottom: 4px;
}
.visual-label { background: rgba(240,128,48,.15); color: var(--orange); }
.text-label   { background: rgba(61,139,255,.15); color: var(--accent); }
.pipe-principles {
  display: flex; flex-wrap: wrap; gap: 10px; justify-content: center;
  padding-top: 20px; border-top: 1px solid var(--border);
}
.principle {
  font-size: 12px; color: var(--text2);
  background: var(--bg3); border-radius: 20px;
  padding: 5px 16px;
}

/* ─── Devlog ─── */
.devlog-group { margin-bottom: 48px; }
.group-header {
  display: flex; align-items: center; justify-content: space-between;
  padding: 14px 20px; border-radius: var(--r);
  font-size: 14px; font-weight: 600; color: var(--text);
  margin-bottom: 8px;
}
.p3-header { background: rgba(61,139,255,.1);  border: 1px solid rgba(61,139,255,.25); }
.sl-header { background: rgba(155,109,255,.1); border: 1px solid rgba(155,109,255,.25); }
.group-count { font-family: var(--mono); font-size: 11px; color: var(--text3); }
.group-desc { font-size: 12px; color: var(--text3); margin-bottom: 20px; padding-left: 4px; line-height: 1.6; }

.exp-card {
  background: var(--bg2); border: 1px solid var(--border);
  border-radius: var(--r); padding: 20px 24px;
  margin-bottom: 14px;
}
.exp-header {
  display: flex; align-items: center; gap: 10px;
  margin-bottom: 16px; flex-wrap: wrap;
}
.exp-id    { font-family: var(--mono); font-size: 13px; font-weight: 500; color: var(--accent); min-width: 42px; }
.exp-title { font-size: 13px; font-weight: 500; color: var(--text); flex: 1; }
.badge {
  display: inline-block; padding: 2px 8px; border-radius: 4px;
  font-size: 10px; font-family: var(--mono); color: #fff; font-weight: 600;
}
.exp-field { margin-bottom: 14px; }
.exp-field:last-child { margin-bottom: 0; }
.exp-field-lbl {
  font-size: 10px; font-weight: 600; text-transform: uppercase;
  letter-spacing: .07em; color: var(--text3); margin-bottom: 5px;
}
.exp-field-val {
  font-size: 13px; color: var(--text2); line-height: 1.75;
  /* Keep text readable — don't collapse into one blob */
  white-space: pre-wrap;
  word-break: break-word;
}

.muted { color: var(--text3); font-size: 12px; font-style: italic; }
"""

JS = """
function showTab(id) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  document.querySelector('.tab[data-tab="'+id+'"]').classList.add('active');
  document.getElementById(id).classList.add('active');
}
"""

# ---------------------------------------------------------------------------
# BUILD
# ---------------------------------------------------------------------------

def build() -> str:
    devlog_md  = read_file(MDS_DIR / "DEVLOG.md")
    devlog     = parse_devlog(devlog_md) if devlog_md else {"pillar3": [], "slides": []}
    summaries  = load_run_summaries()
    audit_dirs = find_audit_dirs()
    gen_ts     = datetime.now().strftime("%Y-%m-%d %H:%M")

    story = f"""
    <div class="section-block">
      <div class="section-title">Pipeline Architecture</div>
      {render_pipeline_diagram()}
    </div>

    <div class="section-block">
      <div class="section-title">Latest Runs &amp; Downloads</div>
      {render_run_panel(summaries) or '<p class="muted">No run summaries found.</p>'}
    </div>

    {render_slide_grid(audit_dirs)}
    """

    devlog_html = render_devlog(devlog) if devlog_md else '<p class="muted">MDs/DEVLOG.md not found.</p>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>FinDocIQ — Project Story</title>
  <style>{CSS}</style>
</head>
<body>
  <header class="site-header">
    <div class="header-top">
      <div>
        <div class="logo">FinDocIQ</div>
        <div class="logo-sub">Financial Document Intelligence — UOB AI Innovation Group</div>
      </div>
      <div class="header-meta">Generated {gen_ts}<br>DBS · OCBC · UOB</div>
    </div>
    <div class="tab-bar">
      <div class="tab active" data-tab="story"  onclick="showTab('story')">📖 Story</div>
      <div class="tab"        data-tab="devlog" onclick="showTab('devlog')">🔬 Dev Log</div>
    </div>
  </header>

  <div id="story"  class="tab-content active">{story}</div>
  <div id="devlog" class="tab-content">{devlog_html}</div>

  <script>{JS}</script>
</body>
</html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out",  default=None)
    ap.add_argument("--open", action="store_true")
    args = ap.parse_args()

    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    html     = build()
    out_path = args.out or str(REPORTS_DIR / "story_report.html")
    Path(out_path).write_text(html, encoding="utf-8")
    size_kb = Path(out_path).stat().st_size / 1024
    print(f"✓ Story report → {out_path}  ({size_kb:.0f} KB)")
    if args.open:
        webbrowser.open(f"file://{Path(out_path).resolve()}")

if __name__ == "__main__":
    main()
