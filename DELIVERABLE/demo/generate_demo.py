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
OUTPUTS_DIR = DELIVERABLE / "outputs" / "slides"
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
    steps = [
        ("01", "Classify", "Pass 0 — micro-call (~150 tokens) identifies every chart type on the slide. Routes the slide to the right extraction branch.", "#8e44ad"),
        ("02", "Single-Pass\n(Visual)", "Charts, donuts, waterfalls → one Gemini call with the image present throughout. No intermediate text. No schema pressure during visual reading.", "#e67e22"),
        ("02", "Multi-Pass\n(Text Tables)", "P&L tables → Pass 1 describes structure and pre-maps every cell to schema fields. Pass 2 transcribes text-only — no re-reading the image.", "#2980b9"),
        ("03", "Validate", "Arithmetic self-checks run on every element. Waterfall bridges must balance. Blank labels are flagged. One correction pass fires if errors are found.", "#27ae60"),
        ("04", "Render", "Wide-format Excel pivot — periods become columns, series become rows. Brand colours, bold totals, yellow chart-sourced cells. One tab per slide.", "#d68910"),
    ]
    cards = ""
    for num, name, desc, colour in steps:
        cards += f"""
        <div class="how-card">
          <div class="how-num" style="color:{colour}">{num}</div>
          <div class="how-name">{name.replace(chr(10),"<br>")}</div>
          <div class="how-desc">{desc}</div>
        </div>"""

    principles = [
        ("No schema pressure during vision", "Visual slides send the image through the entire call. Gemini reads and structures simultaneously — the way it works naturally."),
        ("Contracts teach method, not values", "Each chart type has a reading contract: how to identify periods, what arithmetic must balance, what the colour legend means. No hardcoded expected values."),
        ("Audit trail per slide", "Every slide saves its PNG, the prompt, the raw API response, and parsed datapoints. Resume is free — re-runs skip cached slides."),
        ("Cost stays low", "Pass 0 is ~$0.0001/slide. A full 22-slide deck costs ~$0.18. The entire DBS + OCBC + UOB run across 73 slides costs under $0.50."),
    ]
    pcards = ""
    for title, desc in principles:
        pcards += f"""
        <div class="principle-card">
          <div class="principle-title">{title}</div>
          <div class="principle-desc">{desc}</div>
        </div>"""

    return f"""
    <section class="section" id="how">
      <div class="section-inner">
        <div class="section-eyebrow">Architecture</div>
        <h2 class="section-title">How it works</h2>
        <p class="section-sub">Four passes, one image per slide, under $0.01 per slide.</p>
        <div class="how-grid">{cards}</div>
        <div class="principles-grid">{pcards}</div>
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

/* ── How it works ── */
.how-grid {
  display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
  gap: 20px; margin-bottom: 56px;
}
.how-card {
  background: var(--bg); border: 1px solid var(--border);
  border-radius: var(--r); padding: 28px 24px;
}
.how-num  { font-family: var(--mono); font-size: 13px; font-weight: 600; margin-bottom: 10px; }
.how-name { font-size: 17px; font-weight: 700; color: var(--dark); margin-bottom: 10px; line-height: 1.3; }
.how-desc { font-size: 14px; color: var(--mid); line-height: 1.65; }

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
