"""
build_toc.py — deterministic Table of Contents extraction. ZERO API calls.

Template-agnostic: handles both
  * Part-structured filings (DBS full Pillar 3: "PART A", page labels "A-2"), and
  * Sequentially-numbered filings (OCBC / UOB: "1 Introduction … 5", plain page
    numbers in the footer).

How it anchors sections to physical pages without hard-coded page numbers:
  1. Parse the printed CONTENTS for the section hierarchy and each section's
     printed page reference (a letter label like "A-2" OR a plain number like "5").
  2. Scan every physical page's footer for its printed page token
     ("A-2", "Page 5", or a trailing page number) -> {token: physical_page}.
  3. Map each section's page reference through that table -> physical start page.
  4. (Fallback) if a section can't be anchored that way, use the PDF's embedded
     outline (bookmarks), when present.

Output -> outputs/step1_toc.json   (document, provenance, parts, sections[])

Usage:
  python build_toc.py "DBS_4Q25_Pillar3.pdf"
  python build_toc.py "OCBC_4Q25_Pillar 3.pdf" --out outputs/ocbc_toc.json
"""
from __future__ import annotations
import os, sys, re, json, argparse
from pathlib import Path
import pypdfium2 as pdfium
try:
    import pdfplumber                # better text ordering for footer scans
except Exception:
    pdfplumber = None

TOC_OUT = str(Path(__file__).parent.parent / "outputs" / "pillar3" / "step1_toc.json")

PART_RE  = re.compile(r"^PART\s+([A-Z])\b\s*[:\-]?\s*(.+?)(?:\.{2,}|…|$)", re.I)
# section line WITH a printed page ref (letter label A-2 or plain integer) at the end.
SEC_RE   = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+(.+?)[\s.…]{2,}\s*([A-E]-\d{1,3}|\d{1,4})\s*$")
# section line WITHOUT a page ref (e.g. DBS subsections): a DOTTED number + optional title.
# Title is optional to handle DBS 6.1/6.2/6.3 which have numbers but blank titles.
SEC_NOREF_RE = re.compile(r"^(\d+\.\d+(?:\.\d+)*)\.?\s*([A-Za-z(].*?)?\s*$")
LABEL_RE = re.compile(r"^([A-E])\s*-\s*(\d{1,3})$")           # standalone letter label

def _norm_ref(ref: str) -> str:
    """Normalise a page reference token: 'A - 2' -> 'A-2'; '007' -> '7'."""
    if not ref:
        return ""
    ref = ref.strip()
    m = re.fullmatch(r"([A-E])\s*-\s*(\d{1,3})", ref)
    if m:
        return f"{m.group(1).upper()}-{int(m.group(2))}"
    if ref.isdigit():
        return str(int(ref))
    return ref

# ---------------------------------------------------------------------------
def _page_text(pdf, i: int) -> str:
    return pdf[i].get_textpage().get_text_range()

def _find_contents_start(pdf, max_scan: int = 8) -> int:
    for i in range(min(max_scan, len(pdf))):
        t = _page_text(pdf, i).upper()
        if "CONTENT" in t:                       # "Contents" / "Table of Contents"
            return i
    return 1

_SEC_LINE_RE = re.compile(r"^\s*\d+(?:\.\d+)*\.?\s+\S")   # any "N.N.N  Title" line

def _toc_section_density(txt: str) -> int:
    """Count lines that look like section entries (number + text)."""
    return sum(1 for l in txt.splitlines() if _SEC_LINE_RE.match(l))

def _looks_like_toc(txt: str) -> bool:
    return bool(re.search(r"CONTENT", txt, re.I) or
                re.search(r"^PART\s+[A-Z]", txt, re.M) or
                _toc_section_density(txt) >= 3)

def _collect_contents_text(pdf, start_idx: int) -> str:
    """Collect all TOC pages. Stops only when a page has NO section-number lines
    AND more than 4 long prose lines — i.e. we've genuinely entered body text.
    Page headers/footers ('Pillar 3 Disclosure Report', 'Page 3') are ignored."""
    n = len(pdf)
    out = ""
    for ci in range(start_idx, min(start_idx + 8, n)):
        txt = _page_text(pdf, ci)
        # strip header/footer noise before measuring prose density
        clean_lines = [l for l in txt.split("\n")
                       if l.strip() and not re.match(r"^\s*(Page\s+\d+|Pillar\s+3\b)", l, re.I)]
        clean_txt = "\n".join(clean_lines)
        # long prose lines (after removing dot leaders) signal body text
        prose_lines = [l for l in clean_lines
                       if len(re.sub(r"[.…]{2,}", "", l).strip()) > 130]
        has_sections = _toc_section_density(clean_txt) >= 2
        if ci > start_idx and not has_sections and len(prose_lines) > 4:
            break   # no section entries AND long prose — we've hit body text
        out += "\n" + txt
    return out

SEC_START   = re.compile(r"^\s*\d+(?:\.\d+)*\.?\s+\S")           # line begins a section entry
PAGEREF_END = re.compile(r"([A-E]-\d{1,3}|\d{1,4})\s*$")          # entry ends with a page ref

def _merge_wrapped_lines(text: str) -> list[str]:
    """Collapse wrapped TOC entries into one logical line each. A long title can
    wrap so its page number lands on the next line; we append continuation lines
    to the current entry until it ends with a page reference (or the next
    section/part begins)."""
    logical: list[str] = []
    cur, cur_done = None, False
    for raw in text.replace("\r", "").split("\n"):
        s = raw.strip()
        if not s:
            continue
        sclean = re.sub(r"[\.…]{2,}", "  ", s)
        is_part = bool(PART_RE.match(sclean))
        is_sec  = bool(SEC_START.match(s))
        if is_part or is_sec:
            if cur is not None:
                logical.append(cur)
            cur = s
            cur_done = is_part or bool(PAGEREF_END.search(s))     # parts/single-line entries are complete
        elif cur is not None and not cur_done:
            cur = cur + " " + s                                   # wrapped continuation
            if PAGEREF_END.search(cur):
                cur_done = True
        # else: stray line after a completed entry (footer, etc.) — ignore
    if cur is not None:
        logical.append(cur)
    return logical

_BARE_NUM_RE = re.compile(r"^\s*(\d+\.\d+(?:\.\d+)*)\s*$")  # line is ONLY a section number

def _patch_twocol_toc(raw_text: str) -> str:
    """DBS PDF two-column layout: pypdfium2 emits bare number entries first
    ('6.1', '6.2', '6.3'), then their titles appear as a detached block later.

    For each bare-number entry N.M, we scan forward through the orphan title
    block and read lines until we see the next sibling number N.M+1 (or the block
    ends).  That gives us the full (possibly-wrapped) title for each entry."""
    lines = raw_text.replace("\r", "").split("\n")
    n = len(lines)

    # collect bare-number line indices: line is ONLY "N.M" or "N.M.K" with no title
    bare: dict[int, str] = {}
    for i, l in enumerate(lines):
        m = _BARE_NUM_RE.match(l)
        if m:
            bare[i] = m.group(1)

    if not bare:
        return raw_text

    # group consecutive (within 3 lines) bare-number entries
    bare_indices = sorted(bare)
    groups: list[list[int]] = []
    cur: list[int] = [bare_indices[0]]
    for idx in bare_indices[1:]:
        if idx - cur[-1] <= 3:
            cur.append(idx)
        else:
            groups.append(cur); cur = [idx]
    groups.append(cur)

    consumed: set[int] = set()
    replacements: dict[int, str] = {}

    for grp in groups:
        if len(grp) < 2:
            continue  # single bare number is likely a data cell, not a TOC anomaly

        # build a set of the sibling numbers so we know when one title ends
        sibling_nums = {bare[i] for i in grp}

        # locate the orphan title block: first non-numeric, non-empty line after the group
        start_search = grp[-1] + 1
        orphan_start = None
        for j in range(start_search, min(start_search + 30, n)):
            l = lines[j].strip()
            if not l:
                continue
            sclean = re.sub(r"[\.…]{2,}", "  ", l)
            if SEC_START.match(l) or PART_RE.match(sclean):
                continue  # skip over sibling section entries that have titles already
            orphan_start = j
            break

        if orphan_start is None:
            continue

        # Collect all orphan title lines starting at orphan_start.
        # Stop at any section entry, PART marker, or page header.
        orphan_lines: list[str] = []
        orphan_idxs: list[int] = []
        j = orphan_start
        while j < n:
            l = lines[j].strip()
            if not l:
                j += 1
                continue
            sclean = re.sub(r"[\.…]{2,}", "  ", l)
            if SEC_START.match(l) or PART_RE.match(sclean):
                break
            if re.search(r"CONTENTS\s+Page|DBS GROUP|OCBC|UOB GROUP", l, re.I):
                break
            orphan_lines.append(l)
            orphan_idxs.append(j)
            j += 1

        nslots = len(grp)
        total = len(orphan_lines)
        if total < nslots:
            continue  # not enough lines to fill all slots

        # Divide evenly: each slot gets (total // nslots) lines;
        # the last slot absorbs any remainder.
        per = total // nslots
        for k, bare_idx in enumerate(grp):
            start = k * per
            end = start + per if k < nslots - 1 else total
            title = " ".join(orphan_lines[start:end])
            replacements[bare_idx] = f"{bare[bare_idx]} {title}"
        for idx in orphan_idxs:
            consumed.add(idx)

    result = []
    for i, l in enumerate(lines):
        if i in consumed:
            continue
        result.append(replacements.get(i, l))
    return "\n".join(result)


def _parse_contents(text: str) -> list[dict]:
    """Parse printed CONTENTS into ordered Part + section entries.
    PART headers are optional (DBS has them, OCBC/UOB do not). Wrapped TOC lines
    (long titles whose page number falls on the next line) are merged first."""
    entries, current_part = [], None
    for line in _merge_wrapped_lines(_patch_twocol_toc(text)):
        sclean = re.sub(r"[\.…]{2,}", "  ", line)
        m = PART_RE.match(sclean)
        if m:
            current_part = m.group(1).upper()
            entries.append({"kind": "part", "part": current_part,
                            "title": m.group(2).strip().rstrip(". ")})
            continue
        m = SEC_RE.match(line)                    # 1) section with a printed page ref
        if m:
            number, title, ref = m.group(1), m.group(2), _norm_ref(m.group(3))
        else:                                     # 2) section without a page ref (e.g. DBS subsections)
            m = SEC_NOREF_RE.match(line)
            if not m:
                continue
            number = m.group(1)
            title  = (m.group(2) or "").strip().rstrip(". ") or number  # blank title -> use number
            ref    = ""
        # For part-structured docs (DBS), prefix the number with the part letter
        # so B.1.1 stays B.1.1 rather than colliding with A.1.1
        if current_part:
            sid = f"{current_part}.{number}"
        else:
            sid = number
        entries.append({"kind": "section", "part": current_part, "number": number,
                        "section_id": sid, "title": title, "page_ref": ref})
    # dedupe by id/part, keep first
    seen, out = set(), []
    for e in entries:
        key = f"{e['kind']}:{e.get('section_id') or e.get('part')}"
        if key not in seen:
            seen.add(key); out.append(e)
    return out

def _footer_token(lines: list[str]) -> str:
    """Extract a page token from page header/footer lines.
    Handles:
      - 'A-2' standalone letter label (DBS Part-structured)
      - 'Page 5' anywhere
      - trailing letter label e.g. 'A-2' at end of line
      - OCBC header pattern: 'Pillar 3 Disclosures December 2025 6' (page# at end of first line)
      - short footer with trailing page number
    """
    if not lines:
        return ""
    # OCBC: page number is the last token on the first line, after a year
    # e.g. "Pillar 3 Disclosures December 2025 6"
    first = lines[0]
    m = re.search(r"\b20\d\d\s+(\d{1,4})\s*$", first)
    if m:
        return str(int(m.group(1)))

    # UOB: "Page 36" appears as the second line under a title header
    for l in lines[:3]:
        m = re.match(r"^Page\s+(\d{1,4})\s*$", l, re.I)
        if m:
            return str(int(m.group(1)))

    tail = lines[-2:]
    for l in tail:                                # standalone letter label
        m = LABEL_RE.match(l)
        if m:
            return f"{m.group(1)}-{int(m.group(2))}"
    for l in tail:                                # "Page 5"
        m = re.search(r"\bPage\s+(\d{1,4})\b", l, re.I)
        if m:
            return str(int(m.group(1)))
    for l in tail:                                # trailing letter label
        m = re.search(r"([A-E]-\d{1,3})\s*$", l)
        if m:
            return _norm_ref(m.group(1))
    if tail:                                      # trailing page number on a short footer
        last = tail[-1]
        # reject lines that look like data rows: contain commas, multiple numbers,
        # currency symbols, or are too long — these are table footers not page footers
        if (len(last) <= 60
                and "," not in last
                and "$" not in last
                and len(re.findall(r"\d+", last)) <= 2):
            m = re.search(r"(\d{1,4})\s*$", last)
            if m:
                return str(int(m.group(1)))
    return ""

def _scan_page_map(pdf_path: str, pdf) -> dict[str, int]:
    """token -> physical page (1-based). First occurrence wins.
    Uses pdfplumber for correct visual line order (the footer must be the last
    line); falls back to pypdfium2 if pdfplumber is unavailable."""
    m: dict[str, int] = {}
    if pdfplumber is not None:
        with pdfplumber.open(pdf_path) as pl:
            for i, page in enumerate(pl.pages):
                lines = [l.strip() for l in (page.extract_text() or "").splitlines() if l.strip()]
                tok = _footer_token(lines)
                if tok and tok not in m:
                    m[tok] = i + 1
    else:
        for i in range(len(pdf)):
            lines = [l.strip() for l in _page_text(pdf, i).splitlines() if l.strip()]
            tok = _footer_token(lines)
            if tok and tok not in m:
                m[tok] = i + 1
    return m

def _outline_map(pdf) -> dict[str, int]:
    """Embedded bookmarks -> {leading-number-in-title: physical page}. Fallback only."""
    out: dict[str, int] = {}
    try:
        for b in pdf.get_toc():
            dest = b.get_dest()
            pg = (dest.get_index() + 1) if dest else None
            title = (b.get_title() or "").strip()
            mm = re.match(r"^(\d+(?:\.\d+)*)\b", title)
            if pg and mm and mm.group(1) not in out:
                out[mm.group(1)] = pg
    except Exception:
        pass
    return out

def _words(s: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", (s or "").lower())

def _page_lines(pdf_path: str, pdf) -> list[list[str]]:
    """Per physical page, its non-empty text lines (visual order via pdfplumber)."""
    out: list[list[str]] = []
    if pdfplumber is not None:
        with pdfplumber.open(pdf_path) as pl:
            for page in pl.pages:
                out.append([l.strip() for l in (page.extract_text() or "").splitlines() if l.strip()])
    else:
        for i in range(len(pdf)):
            out.append([l.strip() for l in _page_text(pdf, i).splitlines() if l.strip()])
    return out

def _heading_page(page_lines: list[list[str]], number: str, title: str,
                  start: int, end: int, top_lines: int = 9) -> int | None:
    """First page in [start, end] whose content matches this subsection.

    Pass 1 (precise): ANY line on the page that STARTS with the exact subsection
      number, e.g. "11.2 Comparison of Modelled...". Scans the full page because
      some sections begin mid-page (their heading is not in the first 9 lines).
      The (?!\\d) guard stops "12.2.1" matching "12.2.10".
    Pass 2 (fallback): a line in the top `top_lines` whose words cover >=80% of
      the title — kept narrow to avoid false positives from table data."""
    n = len(page_lines)
    hi = min(end, n)
    num_re = re.compile(rf"^{re.escape(number)}(?!\d)\b")
    for p in range(start, hi + 1):
        for ln in page_lines[p - 1]:          # full page scan for exact number match
            if num_re.match(ln):
                return p
    tw = set(_words(title))
    if tw:
        for p in range(start, hi + 1):
            for ln in page_lines[p - 1][:top_lines]:   # top lines only for fuzzy match
                if len(tw & set(_words(ln))) / len(tw) >= 0.80:
                    return p
    return None

# ---------------------------------------------------------------------------
def build_toc(pdf_path: str) -> dict:
    pdf = pdfium.PdfDocument(pdf_path)
    n = len(pdf)
    warnings: list[str] = []

    c_idx = _find_contents_start(pdf)
    contents = _parse_contents(_collect_contents_text(pdf, c_idx))
    token_map   = _scan_page_map(pdf_path, pdf)
    outline_map = _outline_map(pdf)

    parts = [e for e in contents if e["kind"] == "part"]
    secs  = [e for e in contents if e["kind"] == "section"]
    style = "part_structured" if parts else "numbered"

    # LEAF sections = the smallest subsections (no child whose number extends them).
    # e.g. given 5, 5.1, 5.2 -> leaves are 5.1 and 5.2; a section like 4 with no
    # children is itself a leaf. These become one tab each.
    def _is_leaf(s):
        pre = s["number"] + "."
        return not any(o["number"].startswith(pre) for o in secs if o is not s)
    leaves = [s for s in secs if _is_leaf(s)]
    if not leaves:
        warnings.append("no sections parsed from contents page")

    # Sort key: part letter maps to a number (A=0, B=1, C=2...) so that Part C
    # sections always sort AFTER Part B, which sorts after Part A.
    _PART_ORDER = {None: 0, "A": 0, "B": 1, "C": 2, "D": 3, "E": 4}
    def _numlist(s):
        part_ord = _PART_ORDER.get(s.get("part"), 0)
        return [part_ord] + [int(x) for x in s["number"].split(".")]

    # --- 1) anchor TOP-LEVEL sections first (they carry the printed page refs) ---
    top = [s for s in secs if "." not in s["number"]]
    for s in top:
        phys = token_map.get(s["page_ref"]) if s["page_ref"] else None
        if phys is None:
            phys = outline_map.get(s["number"])
        # part-structured fallback: Part B section 1 -> try token "B-1"
        if phys is None and s.get("part"):
            derived = f"{s['part']}-{s['number']}"
            phys = token_map.get(derived)
        s["start_page"] = phys
    last = None
    for s in sorted(top, key=_numlist):
        if s.get("start_page") is None:
            s["start_page"] = last if last is not None else 1
            warnings.append(f"top-level {s['section_id']} ('{s['title'][:30]}') unanchored; inherited p{s['start_page']}")
        last = s["start_page"]
    top.sort(key=lambda s: (s["start_page"], _numlist(s)))
    # Key by section_id (part-aware) so Part B "1" and Part C "1" don't collide.
    top_range: dict[str, tuple[int, int]] = {}
    for i, s in enumerate(top):
        nxt = top[i + 1]["start_page"] if i + 1 < len(top) else n + 1
        top_range[s["section_id"]] = (s["start_page"], nxt - 1)

    # --- 2) anchor each leaf: page ref -> outline -> HEADING SEARCH (body) -> inherit ---
    # In part-structured docs (DBS), bare numbers like "1.1" appear in multiple
    # parts (B.1.1 AND C.1.1 both have number="1.1"). The outline only has one
    # entry per number so we skip outline lookup for leaves that have a part prefix,
    # to avoid C.1.1 inheriting B.1.1's bookmark.
    def _outline_get(s):
        part = s.get("part")
        if part and part != "A":  # B/C/D/E: section numbers repeat across parts, skip outline
            return None
        return outline_map.get(s["number"])

    # always build page_lines — needed for heading search and end_page refinement
    page_lines = _page_lines(pdf_path, pdf)

    anchored = heading_anchored = 0
    for s in leaves:
        if "." not in s["number"]:                       # top-level leaf already anchored above
            if s.get("start_page"):
                anchored += 1
            continue
        phys = token_map.get(s["page_ref"]) if s["page_ref"] else None
        if phys is None:
            phys = _outline_get(s)
        if phys is None and page_lines is not None:       # deterministic body heading search
            # look up parent by section_id prefix (part-aware: "C.1.1" -> parent "C.1")
            sid_parts = s["section_id"].split(".")
            parent_sid = ".".join(sid_parts[:-1])
            rng = top_range.get(parent_sid) or top_range.get(sid_parts[0])
            if rng:
                phys = _heading_page(page_lines, s["number"], s["title"], rng[0], rng[1])
                if phys:
                    heading_anchored += 1
        s["start_page"] = phys
        if phys:
            anchored += 1

    # fill remaining gaps in document order (heading search missed -> inherit prior leaf)
    last = None
    for s in sorted(leaves, key=_numlist):
        if s.get("start_page") is None:
            s["start_page"] = last if last is not None else 1
            warnings.append(f"{s['section_id']} ('{s['title'][:30]}') "
                            f"unanchored (ref={s['page_ref'] or '—'}); inherited p{s['start_page']}")
        last = s["start_page"]

    def _numkey(s):
        return (s["start_page"], _numlist(s))
    leaves.sort(key=_numkey)
    for i, s in enumerate(leaves):
        nxt = leaves[i + 1]["start_page"] if i + 1 < len(leaves) else n + 1
        s["end_page"] = max(s["start_page"], nxt - 1)

    # Refine end_page using the heading scan: find the last physical page where
    # this section's number appears as a heading (handles both over-extension into
    # trailing content AND under-extension when a section continues onto a shared
    # page that is also the start_page of the next section).
    if page_lines is not None:
        num_re_cache: dict[str, re.Pattern] = {}
        for s in leaves:
            num = s["number"]
            if num not in num_re_cache:
                num_re_cache[num] = re.compile(rf"^{re.escape(num)}(?!\d)\b")
            nr = num_re_cache[num]
            # scan from start_page up to end_page + 1 (one extra to catch cont'd on shared page)
            scan_end = min(s["end_page"] + 1, n)
            last_seen = None
            for p in range(s["start_page"], scan_end + 1):
                for ln in page_lines[p - 1]:
                    if nr.match(ln):
                        last_seen = p
                        break
            if last_seen is not None and last_seen != s["end_page"]:
                s["end_page"] = last_seen

    # flag leaves that share a start page (two subsections on one page -> their
    # tables would otherwise be extracted into both tabs)
    by_start: dict[int, list] = {}
    for s in leaves:
        by_start.setdefault(s["start_page"], []).append(s["section_id"])
    for pg, ids in by_start.items():
        if len(ids) > 1:
            warnings.append(f"subsections share page {pg}: {', '.join(ids)} "
                            f"(tables on that page may need splitting between tabs)")

    part_nodes = []
    for p in parts:
        members = [s for s in leaves if s["part"] == p["part"]]
        if members:
            part_nodes.append({"part": p["part"], "title": p["title"],
                               "start_page": min(s["start_page"] for s in members),
                               "end_page":   max(s["end_page"]   for s in members)})

    title = next((l.strip() for l in _page_text(pdf, 0).splitlines() if l.strip()), "Unknown")

    # build a page-range lookup for intermediate nodes from their leaf descendants
    leaf_by_sid = {s["section_id"]: s for s in leaves}
    def _node_range(number: str, part: str | None) -> tuple[int, int] | None:
        kids = [s for s in leaves if
                s["number"].startswith(number + ".") and s.get("part") == part]
        if not kids:
            return None
        return min(s["start_page"] for s in kids), max(s["end_page"] for s in kids)

    # all_sections: full hierarchy (intermediate + leaf), sorted by document order
    all_sec_list = []
    for s in secs:
        is_leaf = not any(o["number"].startswith(s["number"] + ".") and o.get("part") == s.get("part")
                          for o in secs if o is not s)
        if is_leaf and s["section_id"] in leaf_by_sid:
            ls = leaf_by_sid[s["section_id"]]
            all_sec_list.append({
                "section_id": s["section_id"], "part": s["part"], "number": s["number"],
                "title": s["title"], "page_ref": s["page_ref"],
                "start_page": ls["start_page"], "end_page": ls["end_page"],
                "is_leaf": True,
            })
        else:
            rng = _node_range(s["number"], s.get("part"))
            all_sec_list.append({
                "section_id": s["section_id"], "part": s["part"], "number": s["number"],
                "title": s["title"], "page_ref": s["page_ref"],
                "start_page": rng[0] if rng else None,
                "end_page":   rng[1] if rng else None,
                "is_leaf": False,
            })

    return {
        "document":  {"title": title, "pages": n},
        "provenance": {
            "method": "deterministic: printed contents + footer page-token scan (zero API)",
            "toc_style": style,
            "granularity": "leaf_subsections",
            "contents_page": c_idx + 1,
            "footer_tokens_found": len(token_map),
            "outline_entries": len(outline_map),
            "sections_anchored": f"{anchored}/{len(leaves)}",
            "heading_search_anchored": heading_anchored,
            "warnings": warnings,
        },
        "parts": part_nodes,
        "sections": [                              # LEAVES only — used by extract_to_excel.py
            {"section_id": s["section_id"], "part": s["part"], "number": s["number"],
             "title": s["title"], "page_ref": s["page_ref"],
             "start_page": s["start_page"], "end_page": s["end_page"]}
            for s in leaves
        ],
        "all_sections": all_sec_list,             # FULL hierarchy — for display / audit
    }

# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Deterministic TOC extraction (zero API)")
    ap.add_argument("pdf")
    ap.add_argument("--out", default=TOC_OUT)
    args = ap.parse_args()
    if not os.path.exists(args.pdf):
        sys.exit(f"PDF not found: {args.pdf}")

    toc = build_toc(args.pdf)
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w") as f:
        json.dump(toc, f, indent=2)

    pv = toc["provenance"]
    print(f"📑 {toc['document']['title'][:60]}  ({toc['document']['pages']} pages)")
    print(f"   style={pv['toc_style']}  footer-tokens={pv['footer_tokens_found']}  "
          f"outline={pv['outline_entries']}  anchored={pv['sections_anchored']}")
    print(f"   {len(toc['parts'])} parts, {len(toc['sections'])} top-level sections -> {args.out}")
    for s in toc["sections"]:
        print(f"     {s['section_id']:<8} p{s['start_page']}-{s['end_page']:<3} "
              f"[ref {s['page_ref'] or '—'}]  {s['title'][:46]}")
    if pv["warnings"]:
        print(f"   ⚠️  {len(pv['warnings'])} warning(s):")
        for w in pv["warnings"]:
            print(f"       - {w}")

if __name__ == "__main__":
    main()
