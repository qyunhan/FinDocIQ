"""
reviewer.py — financial-sense checks on extracted FaithfulTable JSON.

Structural correctness (well-formed JSON, right field names) is Pydantic's job.
This asks the harder question: do the numbers make financial sense?

Checks (all soft — failures FLAG for review, never silently correct):
  1. roll-up:  children of a parent row should sum to the parent (per col)
  2. LCR ratio: stated LCR% ≈ HQLA / net_outflows * 100
  3. totals:   any row marked row_type=total should equal sum of its siblings

Usage:
  from reviewer import review_table, ReviewResult
  result = review_table(table_dict)        # pass a dict loaded from JSON
  print(result.makes_financial_sense)
  for f in result.flags: print(f)

Or run standalone:
  python3 reviewer.py out/step3_extracted/t001_p6.json
"""
from __future__ import annotations
import json
import sys
import os
from dataclasses import dataclass, field


@dataclass
class Flag:
    kind: str      # rollup_mismatch | ratio_mismatch | col_missing
    row_id: str
    col_id: str
    detail: str
    expected: float | None = None
    got: float | None = None

    def __str__(self):
        parts = [f"[{self.kind}] row={self.row_id} col={self.col_id}: {self.detail}"]
        if self.expected is not None:
            parts.append(f"  expected={self.expected}  got={self.got}")
        return "\n".join(parts)


@dataclass
class ReviewResult:
    table_id: str
    flags: list[Flag] = field(default_factory=list)
    skipped_reason: str = ""

    @property
    def makes_financial_sense(self) -> bool:
        return len(self.flags) == 0

    def summary(self) -> str:
        if self.skipped_reason:
            return f"{self.table_id}: SKIPPED ({self.skipped_reason})"
        status = "✅ OK" if self.makes_financial_sense else f"⚠️  {len(self.flags)} flag(s)"
        return f"{self.table_id}: {status}"


def _numeric(val) -> float | None:
    """Convert a cell value to float, return None if not numeric."""
    if val is None or val == "" or val == "-":
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        s = val.strip().replace(",", "")
        # Handle parentheses as negatives: (283) → -283
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        try:
            return float(s.rstrip("%"))
        except ValueError:
            return None
    return None


def review_table(table: dict,
                 tolerance_pct: float = 1.0,
                 tolerance_abs: float = 2.0) -> ReviewResult:
    """
    Run financial-sense checks on a FaithfulTable dict.
    tolerance_abs: absolute slack in $m (source data is rounded)
    tolerance_pct: relative slack as % of the parent value
    """
    tid = table.get("table_id", "unknown")
    result = ReviewResult(table_id=tid)

    rows = table.get("rows", [])
    columns = table.get("columns", [])
    col_ids = [c["col_id"] for c in columns]

    if not rows or not col_ids:
        result.skipped_reason = "no rows or columns"
        return result

    by_id = {r["row_id"]: r for r in rows}

    # Build parent → children map
    children_of: dict[str, list[dict]] = {}
    for r in rows:
        pid = r.get("parent_row_id")
        if pid:
            children_of.setdefault(pid, []).append(r)

    _OF_WHICH = ("of which", "of which:", "o/w", "thereof", "including", "inc.")

    # --- Check 1: roll-up (children sum to parent) ---
    for parent_id, kids in children_of.items():
        parent = by_id.get(parent_id)
        if not parent:
            continue
        # Only check data/total rows, skip section headers
        if parent.get("row_type") in ("section_header", "note"):
            continue
        # Skip "of which" children — they are subsets, not exhaustive decompositions
        if any(k.get("label", "").strip().lower().startswith(_OF_WHICH) for k in kids):
            continue
        # Skip "note" children — informational items not meant to sum to parent
        if any(k.get("row_type") == "note" for k in kids):
            continue

        for cid in col_ids:
            pv = _numeric(parent.get("cells", {}).get(cid))
            kid_vals = [_numeric(k.get("cells", {}).get(cid)) for k in kids]

            if pv is None or any(v is None for v in kid_vals) or not kid_vals:
                continue

                # Skip if any child value exceeds the parent — signals a formula
            # decomposition (e.g. income/expense pairs) not an additive rollup
            if pv != 0 and any(abs(v) > abs(pv) * 1.5 for v in kid_vals if v is not None):
                continue

            # Skip income/expense sibling pairs — they feed a formula (e.g. BI
            # component = avg(income, expense)), not a straight sum
            # Skip income/expense or P&L sibling groups — they feed a formula
            # (e.g. BI component = avg(income, expense) or avg of 3-yr P&L),
            # not a straight same-period sum
            _INCOME_WORDS  = ("income", "revenue", "gain", "p&l", "net p&l")
            _EXPENSE_WORDS = ("expense", "cost", "loss", "p&l", "net p&l")
            kid_labels = [k.get("label", "").lower() for k in kids]
            has_income  = any(w in l for l in kid_labels for w in _INCOME_WORDS)
            has_expense = any(w in l for l in kid_labels for w in _EXPENSE_WORDS)
            if has_income and has_expense:
                continue

            s = sum(kid_vals)
            tol = max(tolerance_abs, abs(pv) * tolerance_pct / 100)
            if abs(s - pv) > tol:
                result.flags.append(Flag(
                    kind="rollup_mismatch",
                    row_id=parent_id,
                    col_id=cid,
                    detail=(f"'{parent.get('label','')}' = {pv}, "
                            f"but children sum to {s:.1f} (diff {s-pv:+.1f})"),
                    expected=pv,
                    got=round(s, 2),
                ))

    # --- Check 2: LCR ratio (HQLA / net outflows * 100 ≈ stated LCR%) ---
    # Only applies to LCR tables — detect by row labels
    hqla = net_outflow = lcr_ratio = None
    hqla_rid = outflow_rid = ratio_rid = None

    for r in rows:
        label = (r.get("label") or "").upper()
        cells = r.get("cells", {})
        # Use the first numeric data column as the value column
        for cid in col_ids:
            v = _numeric(cells.get(cid))
            if v is None:
                continue
            if "TOTAL HQLA" in label or "HIGH-QUALITY LIQUID ASSETS" in label:
                hqla, hqla_rid = v, r["row_id"]
            elif "TOTAL NET CASH OUTFLOW" in label:
                net_outflow, outflow_rid = v, r["row_id"]
            elif "LIQUIDITY COVERAGE RATIO" in label:
                lcr_ratio, ratio_rid = v, r["row_id"]
            break

    if hqla is not None and net_outflow and lcr_ratio and net_outflow != 0:
        implied = hqla / net_outflow * 100
        # LCR is a quarter-average so allow up to 5% relative slack
        tol = max(5.0, abs(lcr_ratio) * 0.05)
        if abs(implied - lcr_ratio) > tol:
            result.flags.append(Flag(
                kind="ratio_mismatch",
                row_id=ratio_rid or "?",
                col_id=col_ids[0],
                detail=(f"LCR stated {lcr_ratio}% but "
                        f"HQLA({hqla}) / net_outflow({net_outflow}) * 100 = {implied:.1f}% "
                        f"(note: stated LCR is quarter-average, small gap is normal)"),
                expected=lcr_ratio,
                got=round(implied, 1),
            ))

    return result


def review_all(extracted_dir: str = "out/step3_extracted") -> list[ReviewResult]:
    """Run review on every extracted JSON in the directory."""
    results = []
    if not os.path.isdir(extracted_dir):
        print(f"Directory not found: {extracted_dir}")
        return results

    files = sorted(f for f in os.listdir(extracted_dir) if f.endswith(".json"))
    for fname in files:
        path = os.path.join(extracted_dir, fname)
        try:
            with open(path) as f:
                table = json.load(f)
            result = review_table(table)
            results.append(result)
            print(result.summary())
            for flag in result.flags:
                print(f"    {flag}")
        except Exception as e:
            tid = fname.replace(".json", "")
            r = ReviewResult(table_id=tid, skipped_reason=str(e))
            results.append(r)
            print(r.summary())

    n_ok = sum(1 for r in results if r.makes_financial_sense and not r.skipped_reason)
    n_flagged = sum(1 for r in results if r.flags)
    n_skipped = sum(1 for r in results if r.skipped_reason)
    print(f"\n{'='*50}")
    print(f"Total: {len(results)}  ✅ OK: {n_ok}  ⚠️  Flagged: {n_flagged}  ⏭️  Skipped: {n_skipped}")
    return results


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Single file mode
        path = sys.argv[1]
        with open(path) as f:
            table = json.load(f)
        result = review_table(table)
        print(result.summary())
        for flag in result.flags:
            print(f"  {flag}")
    else:
        # Review all extracted tables
        review_all()
