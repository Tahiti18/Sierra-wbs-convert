# src/roster_enforcer.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
from copy import copy
import csv

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from openpyxl.cell.cell import Cell

# ---------- fixed data paths ----------
ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"        # one name per line, exact top->bottom
ROSTER_CSV = DATA / "gold_master_roster.csv"      # must contain "Employee Name" and "SSN"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"        # gold layout (required for perfect match)

# ---------- load helpers ----------
def _load_order() -> List[str]:
    names = [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines()]
    return [n for n in names if n]

def _load_roster_ssn() -> Dict[str, str]:
    ssn: Dict[str, str] = {}
    with ROSTER_CSV.open("r", encoding="utf-8-sig", newline="") as f:
        rdr = csv.DictReader(f)
        # normalize header names
        header_map = {h: h.strip().lower().replace(" ", "") for h in (rdr.fieldnames or [])}
        name_key = next((h for h, n in header_map.items() if n in ("employeename", "name")), None)
        ssn_key  = next((h for h, n in header_map.items() if n == "ssn"), None)
        if not name_key or not ssn_key:
            return ssn
        for row in rdr:
            nm = (row.get(name_key) or "").strip()
            sv = (row.get(ssn_key)  or "").strip()
            if nm:
                ssn[nm] = sv
    return ssn

# ---------- sheet helpers ----------
def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Return (header_row_1based, column_map) where column_map has keys:
    'Employee Name', 'SSN', 'REGULAR', 'OVERTIME', 'DOUBLETIME', 'Totals' (best-effort).
    """
    colmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 80) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [x.lower().replace(" ", "") for x in labels]
        if any(x in ("employeename",) for x in low):
            for idx, v in enumerate(labels, start=1):
                key = v.strip().lower()
                if key in ("employee name", "employeename"): colmap["Employee Name"] = idx
                elif key == "ssn": colmap["SSN"] = idx
                elif key == "regular": colmap["REGULAR"] = idx
                elif key == "overtime": colmap["OVERTIME"] = idx
                elif key == "doubletime": colmap["DOUBLETIME"] = idx
                elif key == "totals": colmap["Totals"] = idx
            return r, colmap
    raise ValueError("Header row containing 'Employee Name' not found")

def _first_data_row(header_row: int) -> int:
    return header_row + 1

def _collect_rows_by_name(ws: Worksheet, name_col: int, start_row: int) -> Dict[str, int]:
    out: Dict[str, int] = {}
    blanks = 0
    r = start_row
    while r <= ws.max_row and blanks < 15:
        v = ws.cell(row=r, column=name_col).value
        if v is None or str(v).strip() == "":
            blanks += 1
        else:
            blanks = 0
            out[str(v).strip()] = r
        r += 1
    return out

def _copy_style(src: Cell, dst: Cell):
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)
    dst.alignment = copy(src.alignment)

def _clone_row(ws_src: Worksheet, r_src: int, ws_dst: Worksheet, r_dst: int):
    max_col = max(ws_src.max_column, ws_dst.max_column)
    for c in range(1, max_col + 1):
        s = ws_src.cell(row=r_src, column=c)
        d = ws_dst.cell(row=r_dst, column=c)
        d.value = s.value
        _copy_style(s, d)

def _ensure_rows(ws: Worksheet, data_start: int, needed: int, style_row: int):
    """
    Make sure there are at least 'needed' rows available for the data region,
    inserting rows before the footer if required, and styling inserted rows to match style_row.
    """
    # Determine current data capacity: rows from data_start up to just before footer.
    # Footer start heuristic: first index where there are 10 consecutive blank name cells.
    header_row, cmap = _find_header_row(ws)
    name_col = cmap["Employee Name"]
    r = data_start
    blanks = 0
    last_data = data_start - 1
    while r <= ws.max_row and blanks < 10:
        val = ws.cell(row=r, column=name_col).value
        if val is None or str(val).strip() == "":
            blanks += 1
        else:
            blanks = 0
            last_data = r
        r += 1
    footer_start = last_data + 1
    available = max(0, footer_start - data_start)

    if available >= needed:
        return footer_start

    to_insert = needed - available
    # Insert rows right before footer
    ws.insert_rows(footer_start, amount=to_insert)
    # Style newly inserted rows to match style_row
    for i in range(to_insert):
        row_idx = footer_start + i
        for c in range(1, ws.max_column + 1):
            sref = ws.cell(row=style_row, column=c)
            d = ws.cell(row=row_idx, column=c)
            d.value = ""  # values will be filled later
            _copy_style(sref, d)
    # New footer now moves down by 'to_insert'
    return footer_start + to_insert

# ---------- main entry ----------
def enforce_roster(output_wbs_path: str, input_sierra_path: str, _gold_master_path_unused: str) -> None:
    """
    Build final WEEKLY sheet that is a 1:1 match to your gold:
    - order from gold_master_order.txt
    - SSNs from gold_master_roster.csv
    - layout from wbs_template.xlsx
    - values copied from converter output where names match; zeros otherwise
    """
    # Load order + roster SSNs
    order = _load_order()
    ssn_map = _load_roster_ssn()

    # Read the converter's just-produced WBS to harvest weekly values by name
    wb_src: Workbook = load_workbook(output_wbs_path, data_only=False)
    if "WEEKLY" not in wb_src.sheetnames:
        raise ValueError("WEEKLY sheet missing in converter output")
    ws_src: Worksheet = wb_src["WEEKLY"]
    h_src, cmap_src = _find_header_row(ws_src)
    data_start_src = _first_data_row(h_src)
    name_col_src = cmap_src["Employee Name"]
    rows_by_name_src = _collect_rows_by_name(ws_src, name_col_src, data_start_src)

    # Open the gold template
    if not TEMPLATE_XLSX.exists():
        raise FileNotFoundError(f"Missing template: {TEMPLATE_XLSX}")
    wb_dst: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
    if "WEEKLY" not in wb_dst.sheetnames:
        raise ValueError("Template missing 'WEEKLY' sheet")
    ws_dst: Worksheet = wb_dst["WEEKLY"]
    h_dst, cmap_dst = _find_header_row(ws_dst)
    data_start_dst = _first_data_row(h_dst)
    name_col_dst = cmap_dst["Employee Name"]
    ssn_col_dst = cmap_dst.get("SSN")

    # Ensure enough rows in the template to fit the entire order
    style_row = data_start_dst  # first data row as the style source
    _ensure_rows(ws_dst, data_start_dst, len(order), style_row)

    # For each employee in the gold order, write a row
    max_cols = max(ws_src.max_column, ws_dst.max_column)
    for i, emp in enumerate(order):
        r_dst = data_start_dst + i

        # Start from template styling; clear any leftover values (but keep formulas)
        for c in range(1, max_cols + 1):
            cell = ws_dst.cell(row=r_dst, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue
            # simple zero/blank default
            if cell.data_type == 'n':
                cell.value = 0
            else:
                cell.value = ""

        # Name (forced)
        ws_dst.cell(row=r_dst, column=name_col_dst).value = emp

        # SSN (forced from gold CSV)
        if ssn_col_dst is not None:
            ws_dst.cell(row=r_dst, column=ssn_col_dst).value = ssn_map.get(emp, "")

        # If converter produced a row for this name, copy its values into the template row
        src_row = rows_by_name_src.get(emp)
        if src_row:
            for c in range(1, max_cols + 1):
                s = ws_src.cell(row=src_row, column=c)
                d = ws_dst.cell(row=r_dst, column=c)
                # copy the source value; style stays from template
                d.value = s.value
            # re-enforce name + SSN to avoid any drift
            ws_dst.cell(row=r_dst, column=name_col_dst).value = emp
            if ssn_col_dst is not None:
                ws_dst.cell(row=r_dst, column=ssn_col_dst).value = ssn_map.get(emp, "")

    # Save the filled template over the output path (final file)
    wb_dst.save(output_wbs_path)
