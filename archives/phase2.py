"""
phase2.py — docling grid extraction.

For each table in the manifest, runs docling's layout model + TableFormer V2
to produce a precise column grid (with colspan/rowspan) and table bbox.

Outputs per table: out/step2_docling/<table_id>.json
Each file contains:
  {
    "table_id": "t009_p19",
    "page": 19,
    "bbox": {"l": 52.3, "t": 745.4, "r": 572.3, "b": 516.0},
    "n_rows": 15,
    "n_cols": 9,
    "columns": [
      {"col_idx": 0, "header_text": "a", "x_left": 52.3, "x_right": 120.0}
    ],
    "cells": [
      {
        "row": 1, "col": 1,
        "row_span": 1, "col_span": 1,
        "is_col_header": true,
        "is_row_header": false,
        "text": "Closeout uncertainty"
      }
    ]
  }

Usage:
  python3 phase2.py DBS_4Q25_Pillar3.pdf
  python3 phase2.py DBS_4Q25_Pillar3.pdf --table t009_p19
  python3 phase2.py DBS_4Q25_Pillar3.pdf --pages 19,42,66   # specific pages only
"""
from __future__ import annotations
import os
import sys
import csv
import json
import argparse
from collections import defaultdict

# Force CPU before any torch/docling import — Apple MPS doesn't support float64
os.environ["PYTORCH_ENABLE_MPS_FALLBACK"] = "1"
import torch
torch.set_default_device("cpu")

from docling.document_converter import DocumentConverter, PdfFormatOption
from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode
from docling.datamodel.base_models import InputFormat


# ---------------------------------------------------------------------------
# Docling converter — initialised once, reused across all pages
# ---------------------------------------------------------------------------
def make_converter() -> DocumentConverter:
    opts = PdfPipelineOptions()
    opts.do_table_structure = True
    opts.table_structure_options.mode = TableFormerMode.FAST
    opts.do_ocr = False          # text-based PDFs; enable for scanned
    opts.accelerator_options.device = "cpu"
    return DocumentConverter(
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)}
    )


# ---------------------------------------------------------------------------
# Extract docling grid for a single page
# ---------------------------------------------------------------------------
def extract_page_grids(converter: DocumentConverter, pdf_path: str, page_no: int) -> list[dict]:
    """
    Run docling on one page. Returns a list of table grid dicts,
    sorted top-to-bottom by bbox (highest t value first in BOTTOMLEFT coords).
    """
    result = converter.convert(pdf_path, page_range=(page_no, page_no))
    doc = result.document

    grids = []
    for tbl in doc.tables:
        prov = tbl.prov[0] if tbl.prov else None
        if prov is None:
            continue

        bbox = prov.bbox
        data = tbl.data

        # Build column summary from header cells
        col_x: dict[int, dict] = {}
        for cell in data.table_cells:
            ci = cell.start_col_offset_idx
            if ci not in col_x:
                col_x[ci] = {
                    "col_idx":     ci,
                    "header_text": "",
                    "x_left":      None,
                    "x_right":     None,
                }
            if cell.column_header and not col_x[ci]["header_text"]:
                col_x[ci]["header_text"] = cell.text.strip()
            # Track x extents from cell bbox if available
            if cell.bbox:
                if col_x[ci]["x_left"] is None or cell.bbox.l < col_x[ci]["x_left"]:
                    col_x[ci]["x_left"] = round(cell.bbox.l, 2)
                if col_x[ci]["x_right"] is None or cell.bbox.r > col_x[ci]["x_right"]:
                    col_x[ci]["x_right"] = round(cell.bbox.r, 2)

        columns = [col_x[k] for k in sorted(col_x)]

        # Build cell list
        cells = []
        for cell in data.table_cells:
            cells.append({
                "row":           cell.start_row_offset_idx,
                "col":           cell.start_col_offset_idx,
                "row_span":      cell.row_span,
                "col_span":      cell.col_span,
                "is_col_header": cell.column_header,
                "is_row_header": cell.row_header,
                "text":          cell.text.strip(),
                "bbox": {
                    "l": round(cell.bbox.l, 2),
                    "t": round(cell.bbox.t, 2),
                    "r": round(cell.bbox.r, 2),
                    "b": round(cell.bbox.b, 2),
                } if cell.bbox else None,
            })

        grids.append({
            "page":   page_no,
            "bbox":   {
                "l": round(bbox.l, 2),
                "t": round(bbox.t, 2),
                "r": round(bbox.r, 2),
                "b": round(bbox.b, 2),
                "coord_origin": str(bbox.coord_origin),
            },
            "n_rows":   data.num_rows,
            "n_cols":   data.num_cols,
            "columns":  columns,
            "cells":    cells,
        })

    # Sort top-to-bottom: in BOTTOMLEFT coords, higher 't' = higher on page
    grids.sort(key=lambda g: g["bbox"]["t"], reverse=True)
    return grids


# ---------------------------------------------------------------------------
# Match manifest tables to docling grids on same page
# ---------------------------------------------------------------------------
def match_tables_to_grids(
    manifest_rows: list[dict],
    page_grids: dict[int, list[dict]],
) -> list[dict]:
    """
    For each manifest row, find the best-matching docling grid on that page.

    Matching strategy:
    - Single table on page  → take the only grid
    - Multiple tables on page sorted by position:
        * t..._p42 with title "31 Dec 2025" → top grid (highest t)
        * t..._p42 with title "30 Jun 2025" → next grid
      We sort manifest rows for the same page by table_no and pair them
      positionally with docling grids sorted top-to-bottom.
    """
    # Group manifest rows by page
    by_page: dict[int, list[dict]] = defaultdict(list)
    for row in manifest_rows:
        # Use start_page for matching (tables don't span for multiple_on_page)
        page = int(row["start_page"])
        by_page[page].append(row)

    results = []
    for page, rows in sorted(by_page.items()):
        grids = page_grids.get(page, [])
        # Sort manifest rows by table_no to preserve order
        rows_sorted = sorted(rows, key=lambda r: int(r["table_no"]))

        if len(grids) == 0:
            for row in rows_sorted:
                results.append({**row, "docling_grid": None, "match_note": "no docling grids found"})

        elif len(grids) == 1:
            # Only one grid — assign to all manifest rows for this page
            for row in rows_sorted:
                results.append({**row, "docling_grid": grids[0], "match_note": "single grid"})

        else:
            # Multiple grids — pair positionally with manifest rows
            for i, row in enumerate(rows_sorted):
                if i < len(grids):
                    results.append({**row, "docling_grid": grids[i], "match_note": f"grid {i+1}/{len(grids)}"})
                else:
                    # More manifest rows than grids — reuse last grid
                    results.append({**row, "docling_grid": grids[-1], "match_note": f"reused last grid (only {len(grids)} found)"})

    return results


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def run_phase2(pdf_path: str, manifest_path: str, out_dir: str,
               target_table: str | None = None,
               target_pages: list[int] | None = None):

    os.makedirs(out_dir, exist_ok=True)

    with open(manifest_path) as f:
        all_rows = list(csv.DictReader(f))

    # Filter to target if specified
    if target_table:
        rows = [r for r in all_rows if r["table_id"] == target_table]
        if not rows:
            sys.exit(f"table_id '{target_table}' not found in manifest")
    elif target_pages:
        rows = [r for r in all_rows
                if any(int(p) in target_pages for p in r["pages"].split("+"))]
    else:
        rows = all_rows

    # Collect unique pages to process
    pages_needed: set[int] = set()
    for row in rows:
        for p in row["pages"].split("+"):
            pages_needed.add(int(p))

    print(f"  Initialising docling (TableFormer V2, CPU)...")
    converter = make_converter()

    # Process each page once
    page_grids: dict[int, list[dict]] = {}
    total_pages = len(pages_needed)
    for i, page_no in enumerate(sorted(pages_needed), 1):
        print(f"  [{i:>3}/{total_pages}] Page {page_no}...", end="", flush=True)
        try:
            grids = extract_page_grids(converter, pdf_path, page_no)
            page_grids[page_no] = grids
            print(f" {len(grids)} table(s) detected")
        except Exception as e:
            print(f" ❌ {e}")
            page_grids[page_no] = []

    # Match manifest rows to grids
    matched = match_tables_to_grids(rows, page_grids)

    # Save one JSON per table
    saved = 0
    for rec in matched:
        tid = rec["table_id"]
        out_file = os.path.join(out_dir, f"{tid}.json")

        grid = rec["docling_grid"]
        note = rec["match_note"]

        output = {
            "table_id":      tid,
            "section_id":    rec.get("section_id", ""),
            "table_title":   rec.get("table_title", ""),
            "page":          int(rec["start_page"]),
            "pages":         rec["pages"],
            "layout":        rec.get("layout", ""),
            "match_note":    note,
            "docling_grid":  grid,
        }

        with open(out_file, "w") as f:
            json.dump(output, f, indent=2)

        if grid:
            print(f"  ✅ {tid}: {grid['n_rows']}r x {grid['n_cols']}c  [{note}]")
        else:
            print(f"  ⚠️  {tid}: no grid  [{note}]")
        saved += 1

    print(f"\n✅ Phase 2 complete: {saved} tables → {out_dir}/")
    print(f"   Next: python3 extract_tables.py {pdf_path}  (will use docling grids automatically)")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Phase 2: docling layout + TableFormer grid extraction")
    p.add_argument("pdf", help="Path to PDF")
    p.add_argument("--manifest", default="out/step2_table_map.csv")
    p.add_argument("--out",      default="out/step2_docling")
    p.add_argument("--table",    help="Process single table_id only")
    p.add_argument("--pages",    help="Comma-separated page numbers e.g. 19,42,66")
    args = p.parse_args()

    if not os.path.exists(args.pdf):
        sys.exit(f"PDF not found: {args.pdf}")

    target_pages = [int(x) for x in args.pages.split(",")] if args.pages else None

    run_phase2(
        pdf_path=args.pdf,
        manifest_path=args.manifest,
        out_dir=args.out,
        target_table=args.table,
        target_pages=target_pages,
    )
