"""audit_sweep.py — retroactive quality audit over all cached extractions.
No API calls. Grades every parsed.json in the audit cache using validate_numbers.
Run from repo root:
    python3 audit_sweep.py
    python3 audit_sweep.py --bank ocbc
    python3 audit_sweep.py --bank dbs --doc DBS_4Q25_Pillar3
"""
import json, argparse, sys
from pathlib import Path

sys.path.insert(0, "DELIVERABLE/pillar3")
from PASS2_Extract_to_Excel import (
    Extraction, validate_numbers, validate_spans, _normalise_cell_states
)

# Map bank slug → PDF filename in repo root
BANK_PDFS = {
    "dbs":  {"DBS_4Q25_Pillar3": "DBS_4Q25_Pillar3.pdf",
             "DBS_1Q26_Pillar3": "DBS_1Q26_Pillar3.pdf"},
    "ocbc": {"OCBC_4Q25_Pillar 3": "OCBC_4Q25_Pillar 3.pdf",
             "OCBC_1Q26_Pillar3":  "OCBC_1Q26_Pillar3.pdf"},
    "uob":  {"UOB_4Q25_Pillar3": "UOB_4Q25_Pillar 3.pdf",
             "UOB_1Q26_Pillar3": "UOB_1Q26_Pillar3.pdf"},
}

AUDIT_ROOT = Path("DELIVERABLE/outputs/pillar3/audit")

def sweep(bank_filter=None, doc_filter=None):
    results = []

    for bank_dir in sorted(AUDIT_ROOT.iterdir()):
        if not bank_dir.is_dir():
            continue
        bank = bank_dir.name
        if bank_filter and bank != bank_filter:
            continue

        for doc_dir in sorted(bank_dir.iterdir()):
            if not doc_dir.is_dir():
                continue
            doc = doc_dir.name
            if doc_filter and doc != doc_filter:
                continue

            # Resolve PDF path
            pdf_path = BANK_PDFS.get(bank, {}).get(doc)
            if not pdf_path or not Path(pdf_path).exists():
                # try fuzzy match by bank slug
                pdf_path = next(
                    (p for p in Path(".").glob("*.pdf")
                     if bank.upper() in p.name.upper()),
                    None
                )
                if pdf_path:
                    pdf_path = str(pdf_path)

            for unit_dir in sorted(doc_dir.iterdir()):
                if not unit_dir.is_dir():
                    continue
                pj = unit_dir / "parsed.json"
                mj = unit_dir / "meta.json"
                if not pj.exists():
                    continue

                meta = json.load(open(mj)) if mj.exists() else {}
                pages = meta.get("pages", [])
                partial = meta.get("partial", False)
                sids = tuple(meta.get("section_ids", []))

                try:
                    ext = _normalise_cell_states(
                        Extraction(**json.load(open(pj)))
                    )
                except Exception as e:
                    results.append({
                        "bank": bank, "doc": doc, "unit": unit_dir.name,
                        "pages": pages, "n_number": -1, "n_span": -1,
                        "partial": partial, "error": str(e)
                    })
                    continue

                span_issues = validate_spans(ext)
                if pdf_path and pages:
                    try:
                        num_issues = validate_numbers(ext, pdf_path, pages, section_ids=sids)
                    except Exception:
                        num_issues = []
                else:
                    num_issues = []

                results.append({
                    "bank":      bank,
                    "doc":       doc,
                    "unit":      unit_dir.name,
                    "pages":     pages,
                    "n_tables":  len(ext.tables),
                    "n_number":  len(num_issues),
                    "n_span":    len(span_issues),
                    "partial":   partial,
                    "issues":    num_issues[:5],
                })

    return results


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bank", default=None, help="filter by bank (dbs/ocbc/uob)")
    ap.add_argument("--doc",  default=None, help="filter by doc folder name")
    ap.add_argument("--show-clean", action="store_true", help="also print clean units")
    args = ap.parse_args()

    results = sweep(args.bank, args.doc)

    if not results:
        print("No cached audit data found.")
        return

    # Sort: errors first, then by number issues descending
    results.sort(key=lambda x: (-1 if x.get("error") else 0, -x.get("n_number", 0)))

    total = len(results)
    red   = sum(1 for r in results if r.get("n_number", 0) > 10)
    yellow= sum(1 for r in results if 0 < r.get("n_number", 0) <= 10)
    green = sum(1 for r in results if r.get("n_number", 0) == 0 and not r.get("error"))

    print(f"\n{'='*70}")
    print(f"AUDIT SWEEP  —  {total} units  |  🔴 {red}  🟡 {yellow}  🟢 {green}")
    print(f"{'='*70}\n")

    for r in results:
        if r.get("error"):
            print(f"💥 {r['bank']}/{r['unit']:<30} ERROR: {r['error'][:60]}")
            continue

        n   = r["n_number"]
        flag = "🔴" if n > 10 else ("🟡" if n > 0 else "🟢")
        partial_tag = " [PARTIAL]" if r.get("partial") else ""
        span_tag    = f" span:{r['n_span']}" if r["n_span"] else ""
        pages_str   = f"p{r['pages'][0]}-{r['pages'][-1]}" if r.get("pages") else "p?"

        if not args.show_clean and n == 0 and not r.get("partial"):
            continue

        print(f"{flag} {r['bank']}/{r['doc'][-20:]}/{r['unit']:<28} "
              f"{pages_str:<10} {r['n_tables']} tables  {n} issues"
              f"{span_tag}{partial_tag}")

        if n > 0:
            for issue in r.get("issues", [])[:3]:
                print(f"     {issue.strip()}")

    print(f"\nSummary: {total} units — {red} need attention, {yellow} minor, {green} clean")
    if not args.show_clean:
        print("(run with --show-clean to see all units)")


if __name__ == "__main__":
    main()
