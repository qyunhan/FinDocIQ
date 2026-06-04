"""
phase1.py — Steps 1+2: TOC extraction + per-page table mapping.

Strategy (deterministic-first, ported from stage12_map_v3):
  Tier 1 (free, exact): read the PDF's embedded outline for physical page anchors
          + parse the printed contents page for section hierarchy.
          These two sources cross-check each other; neither can hallucinate pages.
  Tier 2 (fallback): if Tier 1 produces implausible results, send only the
          contents-page text to Gemini and ask it to read the structure.
          The LLM is never asked to invent page numbers it can't see.

Step 2 (table map) is still Gemini-driven — it asks Gemini to identify the
individual tables within each section's known page range. This is where
layout/date/title judgment is needed. But it operates on tight, correct page
ranges from Tier 1, not hallucinated ranges.

Produces:
  out/step1_toc.json        — section hierarchy with page ranges
  out/step2_table_map.csv   — one row per table, with table_id, section, title,
                              pages, layout, dates, needs_agent_review

Usage:
  python3 phase1.py DBS_4Q25_Pillar3.pdf
  python3 phase1.py DBS_4Q25_Pillar3.pdf --step1   # TOC only
  python3 phase1.py DBS_4Q25_Pillar3.pdf --step2   # table map only (uses existing TOC)
"""
from __future__ import annotations
import os
import sys
import csv
import json
import re
import time
import argparse
import pypdfium2 as pdfium
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

MODEL      = "gemini-3.5-flash"
MAX_RETRIES = 3

# ===========================================================================
# PYDANTIC SCHEMAS (for Step 2 table-map extraction)
# ===========================================================================

class TableEntry(BaseModel):
    table_title:      str        = Field(default="")
    start_page:       int        = Field(default=0)
    end_page:         int        = Field(default=0)
    layout:           str        = Field(default="single")
    reporting_dates:  list[str]  = Field(default_factory=list)
    notes:            str        = Field(default="")

class SectionEntry(BaseModel):
    section_id:    str              = Field(default="")
    section_title: str              = Field(default="")
    start_page:    int              = Field(default=0)
    end_page:      int              = Field(default=0)
    tables:        list[TableEntry] = Field(default_factory=list)
    subsections:   list["SectionEntry"] = Field(default_factory=list)

SectionEntry.model_rebuild()


# ===========================================================================
# TIER 1 — DETERMINISTIC TOC (outline + contents page)
# ===========================================================================

def _read_outline(pdf) -> list[dict]:
    """PDF embedded bookmarks: exact physical page anchors. Cannot hallucinate."""
    items = []
    try:
        for b in pdf.get_toc():
            try:
                dest = b.get_dest()
                pg = (dest.get_index() + 1) if dest else None
            except Exception:
                pg = None
            items.append({"raw_title": b.get_title(), "level": b.level, "phys_page": pg})
    except Exception:
        pass
    return items


def _find_contents_page(pdf, max_scan: int = 6) -> int:
    """Return the 0-based page index of the printed CONTENTS list."""
    for i in range(min(max_scan, len(pdf))):
        t = pdf[i].get_textpage().get_text_range().upper()
        if "CONTENT" in t and ("PART A" in t or "PART B" in t or "PAGE" in t):
            return i
    return 1


def _parse_contents_page(text: str) -> list[dict]:
    """
    Parse printed CONTENTS into structured entries. Handles:
      PART A : PILLAR 3 DISCLOSURES            → part header
      4 OVERVIEW OF KEY PRUDENTIAL...  A-2     → numbered section + doc-page
      5.1 Financial Statements...              → subsection (no doc-page printed)
    """
    entries = []
    current_part = None

    part_re = re.compile(r"^PART\s+([A-Z])\s*[:\-]\s*(.+?)(?:\.{2,}|…|$)", re.I)
    sec_re  = re.compile(
        r"^(\d+(?:\.\d+)*)\s+(.+?)(?:[\.…\s]{2,}\s*([A-E]-\d+))?\s*$"
    )

    for raw in text.replace("\r", "").split("\n"):
        s = raw.strip()
        if not s:
            continue
        s_clean = re.sub(r"[\.…]{2,}", "  ", s).strip()

        m = part_re.match(s_clean)
        if m:
            part_letter = m.group(1).upper()
            part_title  = m.group(2).strip().rstrip(". ")
            dp = re.search(r"([A-E]-\d+)\s*$", s_clean)
            current_part = part_letter
            entries.append({
                "kind": "part", "part": part_letter,
                "number": part_letter, "title": part_title,
                "doc_page": dp.group(1) if dp else None,
            })
            continue

        m = sec_re.match(s_clean)
        if m and current_part:
            number   = m.group(1)
            title    = m.group(2).strip().rstrip(". ")
            doc_page = m.group(3)
            sec_id   = f"{current_part}.{number}"
            entries.append({
                "kind": "section", "part": current_part,
                "number": number, "section_id": sec_id,
                "title": title, "doc_page": doc_page,
            })

    return entries


def _build_deterministic_toc(pdf_path: str) -> dict:
    """
    Build section hierarchy from the two authoritative sources:
      1. PDF outline → exact physical page anchors (level-0 = part start)
      2. Printed contents page → section titles + hierarchy

    Merges them so every section gets a physical start_page derived from the
    outline (not inferred from text), then estimates end_pages from ordering.
    """
    pdf      = pdfium.PdfDocument(pdf_path)
    n        = len(pdf)
    outline  = _read_outline(pdf)
    cpage_idx = _find_contents_page(pdf)

    # Contents can span multiple consecutive pages — collect them all.
    # A page is still a contents page if it has "PART" or section numbers
    # AND a "CONTENTS Page" header exists within the first 2 pages of the scan.
    combined_text = ""
    CONTENTS_RE = re.compile(r"CONTENTS\s+Page", re.I)
    for ci in range(cpage_idx, min(cpage_idx + 6, n)):
        page_text = pdf[ci].get_textpage().get_text_range()
        is_contents = (
            CONTENTS_RE.search(page_text) or
            (ci == cpage_idx) or
            # continuation: has PART headers or numbered section lines
            bool(re.search(r"^PART\s+[A-E]", page_text, re.M)) or
            bool(re.search(r"^\d+\s+[A-Z]", page_text, re.M))
        )
        # Hard stop: if we see body text (long paragraphs, not a list)
        long_lines = [l for l in page_text.split("\n") if len(l.strip()) > 120]
        if ci > cpage_idx and len(long_lines) > 3:
            break
        if not is_contents and ci > cpage_idx:
            break
        combined_text += "\n" + page_text
    contents  = _parse_contents_page(combined_text)
    # Deduplicate: keep first occurrence of each section_id
    seen_ids: set[str] = set()
    deduped: list[dict] = []
    for e in contents:
        key = f"{e['kind']}:{e.get('section_id') or e.get('part')}"
        if key not in seen_ids:
            seen_ids.add(key)
            deduped.append(e)
    contents = deduped

    structural_re = re.compile(r"^(cover|content)s?$", re.I)

    parts_in_contents    = [e for e in contents if e["kind"] == "part"]
    sections_in_contents = [e for e in contents if e["kind"] == "section"]

    # All non-structural level-1 page anchors from the outline, in order.
    # These are the actual data-page anchors regardless of how DBS labeled them.
    all_anchors = sorted({
        o["phys_page"] for o in outline
        if o["phys_page"] and not structural_re.match((o["raw_title"] or "").strip())
    })

    # Assign Part page ranges by matching known Part keywords to outline titles.
    # DBS outline titles use internal codes, so we match by known Part start pages
    # embedded in the outline: "20 LCR" → Part B, "21 NSFR" → Part C, etc.
    # Match Part B/C/D/E start pages from the outline.
    # Anchor patterns to start-of-title to avoid false matches on internal codes.
    PART_PATTERNS = {
        "B": re.compile(r"^20\s|^Part\s*B\b", re.I),
        "C": re.compile(r"^21\s|^Part\s*C\b", re.I),
        "D": re.compile(r"^22\s|^Part\s*D\b|^Attest", re.I),
        "E": re.compile(r"^23\s|^Part\s*E\b|^Abbrev", re.I),
    }
    part_page_map: dict[str, int] = {}  # part letter → physical start page
    for o in outline:
        if not o.get("phys_page"):
            continue
        title = (o.get("raw_title") or "").strip()
        for letter, pat in PART_PATTERNS.items():
            if pat.match(title):
                if letter not in part_page_map:
                    part_page_map[letter] = o["phys_page"]

    # Part A starts at the first non-structural page after cover/contents
    first_content_anchor = next(
        (o["phys_page"] for o in outline
         if o["phys_page"] and o["phys_page"] >= 5
         and not structural_re.match((o["raw_title"] or "").strip())),
        5
    )
    part_page_map["A"] = first_content_anchor

    # Sort parts by their start page
    ordered_parts = sorted(
        [(p["part"], part_page_map.get(p["part"], 0)) for p in parts_in_contents],
        key=lambda x: x[1]
    )

    for idx, (letter, start) in enumerate(ordered_parts):
        part = next((p for p in parts_in_contents if p["part"] == letter), None)
        if not part:
            continue
        nxt_start = ordered_parts[idx + 1][1] if idx + 1 < len(ordered_parts) else n + 1
        part["start_page"] = start or 1
        part["end_page"]   = nxt_start - 1 if nxt_start <= n else n
        # Anchors for this part = all anchors within [start, nxt_start)
        part["content_anchors"] = [a for a in all_anchors if start <= a < nxt_start]

    # Group sections by part
    by_part: dict[str, list] = {}
    for e in sections_in_contents:
        by_part.setdefault(e["part"], []).append(e)

    # Distribute content anchors across top-level sections within each Part
    for part in parts_in_contents:
        secs     = by_part.get(part["part"], [])
        anchors  = part.get("content_anchors", [])
        top_secs = [s for s in secs if "." not in s["number"]]
        sub_secs = [s for s in secs if "." in s["number"]]

        for i, s in enumerate(top_secs):
            s["start_page"] = anchors[i] if i < len(anchors) else (
                top_secs[i - 1].get("start_page") if i > 0 else part.get("start_page")
            )

        for i, s in enumerate(top_secs):
            nxt_sp = top_secs[i + 1].get("start_page") if i + 1 < len(top_secs) else None
            s["end_page"] = max(s["start_page"], nxt_sp - 1) if nxt_sp else part.get("end_page")

        # Subsections: use free anchors not consumed by top-level, else inherit parent
        used         = {s.get("start_page") for s in top_secs}
        free_anchors = [a for a in anchors if a not in used]
        if sub_secs and len(free_anchors) >= len(sub_secs):
            for s, a in zip(sub_secs, free_anchors):
                s["start_page"] = a
                s["end_page"]   = a
        else:
            for s in sub_secs:
                parent = next(
                    (t for t in top_secs if t["number"] == s["number"].split(".")[0]), None
                )
                if parent:
                    s["start_page"] = parent.get("start_page")
                    s["end_page"]   = parent.get("end_page")

    # Assemble output in the same shape downstream expects
    out_sections = []
    for part in parts_in_contents:
        # Part-level node (no tables at this level — tables live in sections)
        part_secs = by_part.get(part["part"], [])
        top_secs  = [s for s in part_secs if "." not in s["number"]]
        sub_secs  = [s for s in part_secs if "." in s["number"]]

        subsections = []
        for ts in top_secs:
            children = [s for s in sub_secs if s["number"].startswith(ts["number"] + ".")]
            subsections.append({
                "section_id":    ts["section_id"],
                "section_title": ts["title"],
                "start_page":    ts.get("start_page") or part.get("start_page", 0),
                "end_page":      ts.get("end_page")   or part.get("end_page", 0),
                "tables":        [],   # filled in by Step 2
                "subsections": [{
                    "section_id":    c["section_id"],
                    "section_title": c["title"],
                    "start_page":    c.get("start_page") or ts.get("start_page", 0),
                    "end_page":      c.get("end_page")   or ts.get("end_page", 0),
                    "tables":        [],
                    "subsections":   [],
                } for c in children],
            })

        out_sections.append({
            "section_id":    part["part"],
            "section_title": part["title"],
            "start_page":    part.get("start_page", 0),
            "end_page":      part.get("end_page", 0),
            "tables":        [],
            "subsections":   subsections,
        })

    return {
        "document_title": "DBS Group Holdings Ltd Pillar 3 Disclosures",
        "bank_name":      "DBS",
        "filing_type":    "Pillar 3",
        "period":         "Unknown",
        "doc_meta": {
            "pages":           n,
            "contents_page":   cpage_idx + 1,
            "outline_entries": len(outline),
            "method":          "embedded_outline + contents_page (deterministic)",
        },
        "sections": out_sections,
    }


def _toc_is_plausible(toc: dict) -> tuple[bool, str]:
    """Quick sanity checks. If any fail, fall back to Gemini."""
    secs = []
    def collect(sections):
        for s in sections:
            if not s.get("is_part", False):
                secs.append(s)
            collect(s.get("subsections", []))
    collect(toc.get("sections", []))

    if len(secs) == 0:
        return False, "no sections parsed from contents page"
    if len(secs) > 300:
        return False, f"{len(secs)} sections — over-matched"

    n = toc.get("doc_meta", {}).get("pages", 10**9)
    bad = [s for s in secs if not s.get("start_page") or not (1 <= s["start_page"] <= n)]
    if len(bad) > max(3, len(secs) // 3):
        return False, f"{len(bad)}/{len(secs)} sections missing valid page anchor"

    return True, "ok"


def _toc_via_gemini_fallback(pdf_path: str, client: genai.Client) -> dict:
    """
    Fallback: send only the contents-page text to Gemini. It reads the structure
    from what's printed — it is NOT asked to invent page numbers beyond what's shown.
    """
    pdf      = pdfium.PdfDocument(pdf_path)
    n        = len(pdf)
    cidx     = _find_contents_page(pdf)
    contents_text = pdf[cidx].get_textpage().get_text_range()
    outline  = _read_outline(pdf)
    outline_hint = "\n".join(
        f"  page {o['phys_page']}: {o['raw_title']}"
        for o in outline if o.get("phys_page")
    )

    prompt = (
        "Extract the section list from this financial document's contents page.\n"
        "Return a JSON object with key 'sections', an array where each entry has:\n"
        "  section_id (dot-notation, e.g. 'A.12.2.1'), title (verbatim), "
        "start_page (integer, from the embedded outline below if the contents page "
        "doesn't show a physical page), end_page (same as start_page if unknown), "
        "likely_has_tables (true/false), is_part (true only for PART A/B/C/D/E headers).\n"
        "Do NOT invent sections or pages not supported by one of these two sources.\n\n"
        "=== EMBEDDED OUTLINE (physical page → bookmark title) ===\n"
        + (outline_hint or "(none)") +
        "\n\n=== CONTENTS PAGE TEXT ===\n" + contents_text[:8000]
    )

    try:
        resp = client.models.generate_content(
            model=MODEL, contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json", temperature=0.0,
            ),
        )
        raw  = resp.text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1].lstrip("json").strip()
        data = json.loads(raw)
        secs = []
        for s in data.get("sections", []):
            sp = s.get("start_page") or 0
            ep = s.get("end_page") or sp
            secs.append({
                "section_id":    s.get("section_id", ""),
                "section_title": s.get("title", ""),
                "start_page":    sp,
                "end_page":      ep,
                "tables":        [],
                "subsections":   [],
            })
        return {
            "document_title": "Unknown", "bank_name": "Unknown",
            "filing_type": "Pillar 3", "period": "Unknown",
            "doc_meta": {
                "pages": n, "contents_page": cidx + 1,
                "outline_entries": len(outline),
                "method": "agent_fallback (contents-page text)",
            },
            "sections": secs,
        }
    except Exception as e:
        return {
            "document_title": "Unknown", "bank_name": "Unknown",
            "filing_type": "Pillar 3", "period": "Unknown",
            "doc_meta": {
                "pages": n, "contents_page": cidx + 1,
                "outline_entries": len(outline),
                "method": f"agent_fallback ERROR: {e}",
            },
            "sections": [],
        }


def run_step1(pdf_path: str, client: genai.Client, out_path: str) -> dict:
    """
    Tier 1: deterministic outline+contents parse (free, exact).
    Tier 2: Gemini reads only the contents-page text (fallback).
    """
    print("  🔍 Tier 1: deterministic outline + contents-page parse...")
    toc = _build_deterministic_toc(pdf_path)
    ok, reason = _toc_is_plausible(toc)

    if ok:
        toc["doc_meta"]["tier"] = "1_deterministic"
        print(f"  ✅ Tier 1 OK — {toc['doc_meta']['method']}")
        print(f"     {len(toc['sections'])} parts, page anchors from PDF outline (cannot hallucinate)")
    else:
        print(f"  ⚠️  Tier 1 implausible ({reason}) — falling back to Gemini...")
        toc = _toc_via_gemini_fallback(pdf_path, client)
        toc["doc_meta"]["tier"] = "2_agent_fallback"
        toc["doc_meta"]["fallback_reason"] = reason
        print(f"  ⚠️  Tier 2 used — eyeball {out_path} before proceeding!")

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w") as f:
        json.dump(toc, f, indent=2)
    print(f"  → {out_path}")
    return toc


# ===========================================================================
# STEP 2 — GEMINI TABLE MAP (within known page ranges)
# ===========================================================================

STEP2_PROMPT = """You are identifying financial data tables within a specific section of a bank's regulatory disclosure document.

SECTION: {section_id} — {section_title}
PHYSICAL PAGES: {start_page}–{end_page}  (these are exact, from the PDF outline — do not change them)

PAGE TEXT:
{text}

YOUR TASK:
List every distinct data table (grid of financial data) within this section.
Tables in adjacent sections on the same page belong to THOSE sections — do not include them here.

RULES:
1. A "table" is a grid with rows and columns of financial data. Not narrative paragraphs.
2. For each table output: table_title, start_page, end_page, layout, reporting_dates (array), notes.
3. start_page and end_page MUST be within {start_page}–{end_page}. Never output pages outside this range.
4. layout: "single" (one grid, one page), "multiple_on_page" (2+ separate grids on same page),
   "spans_pages" (one continuous grid across pages), "multiple_spanning" (wide grid split by column width).
5. reporting_dates: ALL date headers visible in the table as an array e.g. ["31 Dec 2025", "30 Jun 2025"].
6. If there are no data tables in this section, return an empty array.
7. table_title: use the nearest label printed above the grid. Include date if present.

Return ONLY JSON: {"tables": [...]}
Each table: {"table_title":"...","start_page":N,"end_page":N,"layout":"...","reporting_dates":[...],"notes":""}
"""

def _extract_page_text(pdf_path: str, start_page: int, end_page: int) -> str:
    pdf = pdfium.PdfDocument(pdf_path)
    parts = []
    for pg in range(start_page, min(end_page + 1, len(pdf) + 1)):
        text = pdf[pg - 1].get_textpage().get_text_range()
        parts.append(f"[PAGE {pg}]\n{text}")
    return "\n\n".join(parts)


def _gemini_tables_for_section(
    client: genai.Client,
    pdf_path: str,
    section: dict,
) -> list[dict]:
    """Ask Gemini to find tables within one section's known page range."""
    sp = section.get("start_page", 0)
    ep = section.get("end_page", 0)
    if not sp or not ep or ep < sp:
        return []

    text = _extract_page_text(pdf_path, sp, ep)
    prompt = (STEP2_PROMPT
              .replace("{section_id}",    section.get("section_id", ""))
              .replace("{section_title}", section.get("section_title", ""))
              .replace("{start_page}",    str(sp))
              .replace("{end_page}",      str(ep))
              .replace("{text}",          text))

    for attempt in range(MAX_RETRIES):
        try:
            resp = client.models.generate_content(
                model=MODEL, contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0,
                    max_output_tokens=4096,
                ),
            )
            raw = resp.text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1].lstrip("json").strip()
            data = json.loads(raw)
            tables = data.get("tables", [])
            # Clamp pages to section range — model must not exceed them
            for t in tables:
                t["start_page"] = max(sp, min(ep, int(t.get("start_page") or sp)))
                t["end_page"]   = max(sp, min(ep, int(t.get("end_page") or ep)))
                if isinstance(t.get("reporting_dates"), str):
                    t["reporting_dates"] = [t["reporting_dates"]] if t["reporting_dates"] else []
            return tables
        except Exception as e:
            err = str(e)
            if ("429" in err or "503" in err) and attempt < MAX_RETRIES - 1:
                wait = 2 ** attempt + 1
                print(f"    ⏳ Retry in {wait}s...")
                time.sleep(wait)
            else:
                print(f"    ⚠️  Section {section.get('section_id')} table detection failed: {e}")
                return []
    return []


def _collect_leaf_sections(sections: list[dict]) -> list[dict]:
    """Return sections that have no subsections (leaf nodes) — these hold the tables."""
    leaves = []
    for s in sections:
        subs = s.get("subsections", [])
        if subs:
            leaves.extend(_collect_leaf_sections(subs))
        else:
            leaves.append(s)
    return leaves


def _cap_table_page_ranges(toc_dict: dict) -> dict:
    """Cap each table's end_page using the next table's start_page. Structural, not heuristic."""
    all_tables: list[dict] = []

    def collect(sections: list):
        for s in sections:
            all_tables.extend(s.get("tables", []))
            collect(s.get("subsections", []))

    collect(toc_dict.get("sections", []))

    for i, t in enumerate(all_tables):
        sp = t.get("start_page", 0)
        ep = t.get("end_page", 0)
        if not sp or not ep:
            continue
        if i + 1 < len(all_tables):
            next_sp = all_tables[i + 1].get("start_page", 0)
            if next_sp > 0 and ep >= next_sp:
                new_ep = max(sp, next_sp - 1)
                if new_ep != ep:
                    print(f"    📐 Capped '{t.get('table_title','')[:40]}': end_page {ep}→{new_ep}")
                    t["end_page"] = new_ep
                    if new_ep == sp:
                        t["layout"] = "single"

    return toc_dict


def run_step2_tables(toc: dict, pdf_path: str, client: genai.Client,
                     out_csv: str, out_json: str):
    """
    Walk each leaf section in the TOC. For sections that likely have tables,
    ask Gemini to identify tables within their known (exact) page range.
    Write the flat manifest to CSV + JSON.
    """
    TABLE_WORDS     = re.compile(
        r"METRICS|RATIO|RWA|RISK.WEIGHTED|FLOW STATEMENT|LCR|LEVERAGE|"
        r"OVERVIEW|AVERAGE|COMPARISON|CREDIT RISK|COUNTERPARTY|SECURITI|"
        r"MARKET RISK|CVA|OPERATIONAL|IRRBB|NSFR|BACKTESTING|COLLATERAL|"
        r"EXPOSURE|CAPITAL|COMPOSITION|LINKAGE|PRUDENT|G.SIB|ASSET ENCUMB|"
        r"QUANTITATIVE|DISCLOSURES|LOSSES|INDICATOR|MINIMUM", re.I)
    NARRATIVE_WORDS = re.compile(
        r"^(INTRODUCTION|SCOPE|ABBREVIATION|ATTESTATION|QUALITATIVE\s+DISC)", re.I)

    leaves = _collect_leaf_sections(toc.get("sections", []))
    print(f"  📋 {len(leaves)} leaf sections to scan for tables...")

    for s in leaves:
        title = s.get("section_title", "")
        has_table_signal = (
            bool(TABLE_WORDS.search(title)) and
            not bool(NARRATIVE_WORDS.match(title.strip()))
        )
        sp = s.get("start_page", 0)
        ep = s.get("end_page", 0)

        if not has_table_signal or not sp or not ep:
            s["tables"] = []
            continue

        print(f"  🔍 {s['section_id']} [{sp}–{ep}] {title[:50]}")
        tables = _gemini_tables_for_section(client, pdf_path, s)
        s["tables"] = tables
        if tables:
            print(f"      → {len(tables)} table(s) found")

    # Cap page ranges using ordering constraint
    _cap_table_page_ranges(toc)

    # Flatten to manifest
    records = flatten_toc_to_table_map(toc)

    os.makedirs(os.path.dirname(out_csv) or ".", exist_ok=True)
    if records:
        with open(out_csv, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
            writer.writeheader()
            writer.writerows(records)
    with open(out_json, "w") as f:
        json.dump(records, f, indent=2)

    needs_review = sum(1 for r in records if r["needs_agent_review"] == "True")
    print(f"  ✅ Table map: {len(records)} tables ({needs_review} need review) → {out_csv}")
    return records


# ===========================================================================
# STEP 2 — FLATTEN TOC TO TABLE MAP (unchanged — good as-is)
# ===========================================================================

def flatten_toc_to_table_map(toc: dict) -> list[dict]:
    records  = []
    table_no = [0]

    def walk(section: dict):
        sid    = section.get("section_id", "")
        stitle = section.get("section_title", "")

        for t in section.get("tables", []):
            dates  = t.get("reporting_dates", [])
            layout = t.get("layout", "single")

            if layout == "multiple_on_page" and len(dates) > 1:
                for date in dates:
                    table_no[0] += 1
                    pages_str = "+".join(str(p) for p in range(t["start_page"], t["end_page"] + 1))
                    tid = f"t{table_no[0]:03d}_p{pages_str.replace('+','_')}"
                    records.append({
                        "table_no":      table_no[0],
                        "table_id":      tid,
                        "section_id":    sid,
                        "section_title": stitle,
                        "table_title":   f"{t['table_title']} - {date}",
                        "start_page":    t["start_page"],
                        "end_page":      t["end_page"],
                        "pages":         pages_str,
                        "n_pages":       t["end_page"] - t["start_page"] + 1,
                        "layout":        layout,
                        "dates":         date,
                        "needs_agent_review": str(t.get("notes", "") != ""),
                    })
            else:
                table_no[0] += 1
                pages_str = "+".join(str(p) for p in range(t["start_page"], t["end_page"] + 1))
                tid = f"t{table_no[0]:03d}_p{pages_str.replace('+','_')}"
                records.append({
                    "table_no":      table_no[0],
                    "table_id":      tid,
                    "section_id":    sid,
                    "section_title": stitle,
                    "table_title":   t["table_title"],
                    "start_page":    t["start_page"],
                    "end_page":      t["end_page"],
                    "pages":         pages_str,
                    "n_pages":       t["end_page"] - t["start_page"] + 1,
                    "layout":        layout,
                    "dates":         ", ".join(dates),
                    "needs_agent_review": str(t.get("notes", "") != ""),
                })

        for sub in section.get("subsections", []):
            walk(sub)

    for section in toc.get("sections", []):
        walk(section)

    return records


# keep old run_step2 signature for orchestrator compatibility
def run_step2(toc: dict, out_csv: str, out_json: str):
    """Compatibility shim — flattens an already-populated TOC to CSV/JSON."""
    records = flatten_toc_to_table_map(toc)
    os.makedirs(os.path.dirname(out_csv) or ".", exist_ok=True)
    if records:
        with open(out_csv, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
            writer.writeheader()
            writer.writerows(records)
    with open(out_json, "w") as f:
        json.dump(records, f, indent=2)
    needs_review = sum(1 for r in records if r["needs_agent_review"] == "True")
    print(f"  ✅ Table map: {len(records)} tables ({needs_review} need review) → {out_csv}")
    return records


# ===========================================================================
# CLI
# ===========================================================================

def main():
    p = argparse.ArgumentParser(description="Phase 1: TOC extraction + table mapping")
    p.add_argument("pdf",      help="Path to the PDF file")
    p.add_argument("--step1",  action="store_true", help="TOC only (deterministic)")
    p.add_argument("--step2",  action="store_true", help="Table map only (uses existing TOC)")
    p.add_argument("--toc",   default="out/step1_toc.json",       help="TOC output path")
    p.add_argument("--csv",   default="out/step2_table_map.csv",  help="Table map CSV output")
    p.add_argument("--json",  default="out/step2_table_map.json", help="Table map JSON output")
    args = p.parse_args()

    if not os.path.exists(args.pdf):
        sys.exit(f"PDF not found: {args.pdf}")

    run_both = not args.step1 and not args.step2
    client   = genai.Client() if (args.step2 or run_both) else None

    # Step 1 — deterministic TOC
    if args.step1 or run_both:
        print("🗂  Step 1: TOC (deterministic outline + contents page)...")
        toc = run_step1(args.pdf, client, args.toc)
    else:
        if not os.path.exists(args.toc):
            sys.exit(f"TOC not found: {args.toc}. Run step 1 first.")
        with open(args.toc) as f:
            toc = json.load(f)
        print(f"📂 Using existing TOC: {args.toc}")

    # Step 2 — Gemini finds tables within known page ranges
    if args.step2 or run_both:
        if client is None:
            client = genai.Client()
        print("📋 Step 2: Finding tables within each section (Gemini, tight page ranges)...")
        run_step2_tables(toc, args.pdf, client, args.csv, args.json)

    print("\n✅ Phase 1 complete.")
    print(f"   Eyeball {args.toc} and {args.csv} before running phase2.")
    print(f"   Next: python3 phase2.py {args.pdf}")


if __name__ == "__main__":
    main()
