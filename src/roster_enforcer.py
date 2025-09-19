# src/roster_enforcer.py
from __future__ import annotations
from typing import Dict, List, Tuple
from copy import copy
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from openpyxl.cell.cell import Cell

from gold_roster import load_order, load_roster_csv, ssn_for, template_path

# ---------- Sierra helper (only to grab weekly values if needed) ----------
def _read_sierra_for_names(input_sierra_path: str) -> pd.DataFrame:
    """
    Returns a normalized DF with at least columns: Name, REGULAR, OVERTIME, DOUBLETIME, Totals (best-effort).
    We DON'T use it to decide order — only to confirm values if your WBS row lacks them.
    """
    try:
        df = pd.read_excel(input_sierra_path, sheet_name="WEEKLY", header=7)
        df = df.dropna(how="all")
        # common duplicated header row
        if (df.iloc[0:1].astype(str).apply(lambda x: (x == 'Employee Name').any(), axis=1).any()):
            df = df.iloc[1:]
        name_col = "Employee Name" if "Employee Name" in df.columns else ("Unnamed: 2" if "Unnamed: 2" in df.columns else None)
        if not name_col:
            return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME","Totals"])
        out = pd.DataFrame({
            "Name": df[name_col].astype(str).str.strip()
        })
        for col in ["REGULAR","OVERTIME","DOUBLETIME","Totals"]:
            if col in df.columns:
                out[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                out[col] = 0
        out = out[out["Name"].str.len() > 0]
        return out
    except Exception:
        return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME","Totals"])

# ---------- openpyxl utilities ----------
def _find_header_row(ws: Worksheet) -> Tuple[int, dict]:
    wanted = ["SSN","Employee Name","REGULAR","OVERTIME","DOUBLETIME","Totals"]
    col_idx = {}
    for r in range(1, min(ws.max_row, 60) + 1):
        row_vals = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        if any(v.lower() == "employeename" or v.lower()=="employee name" for v in row_vals):
            for c_idx, v in enumerate(row_vals, start=1):
                v_low = v.strip().lower()
                if v_low in ("ssn",): col_idx["SSN"] = c_idx
                if v_low in ("employee name","employeename"): col_idx["Employee Name"] = c_idx
                if v_low == "regular": col_idx["REGULAR"] = c_idx
                if v_low == "overtime": col_idx["OVERTIME"] = c_idx
                if v_low == "doubletime": col_idx["DOUBLETIME"] = c_idx
                if v_low == "totals": col_idx["Totals"] = c_idx
            return r, col_idx
    raise ValueError("Header row with 'Employee Name' not found")

def _first_data_row(header_row: int) -> int:
    return header_row + 1

def _collect_rows_by_name(ws: Worksheet, name_col: int, start_row: int) -> Dict[str, int]:
    result = {}
    blanks = 0
    r = start_row
    while r <= ws.max_row and blanks < 12:
        name = ws.cell(row=r, column=name_col).value
        if name is None or str(name).strip() == "":
            blanks += 1
        else:
            blanks = 0
            result[str(name).strip()] = r
        r += 1
    return result

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

# ---------- main ----------
def enforce_roster(output_wbs_path: str, input_sierra_path: str, gold_master_path: str) -> None:
    """
    Hard-lock final WEEKLY sheet to:
      - exact gold order (gold_master_order.txt),
      - SSNs from gold_master_roster.csv,
      - preserve layout (prefer data/wbs_template.xlsx if present),
      - copy values for matched names; insert zero rows for missing,
      - keep totals/footer/formulas.
    """
    order = load_order()
    roster_csv = load_roster_csv()
    sierra_df = _read_sierra_for_names(input_sierra_path)

    # Source values workbook (the one converter just produced)
    wb_src: Workbook = load_workbook(output_wbs_path)
    if "WEEKLY" not in wb_src.sheetnames:
        raise ValueError("WEEKLY sheet missing in output WBS")
    ws_src: Worksheet = wb_src["WEEKLY"]

    # Destination workbook: template (prefer) or same workbook
    tpath = template_path()
    if tpath:
        wb_dst: Workbook = load_workbook(tpath)
        if "WEEKLY" not in wb_dst.sheetnames:
            raise ValueError("Template missing 'WEEKLY' sheet")
        ws_dst: Worksheet = wb_dst["WEEKLY"]
    else:
        wb_dst = load_workbook(output_wbs_path)
        if "WEEKLY_enforced" in wb_dst.sheetnames:
            del wb_dst["WEEKLY_enforced"]
        ws_dst = wb_dst.create_sheet("WEEKLY_enforced")

    # Analyze src/dst columns
    h_src, cmap_src = _find_header_row(ws_src)
    h_dst, cmap_dst = _find_header_row(ws_dst)
    name_src, ssn_src = cmap_src.get("Employee Name"), cmap_src.get("SSN")
    name_dst, ssn_dst = cmap_dst.get("Employee Name"), cmap_dst.get("SSN")
    data_start_src, data_start_dst = _first_data_row(h_src), _first_data_row(h_dst)
    rows_by_name = _collect_rows_by_name(ws_src, name_src, data_start_src)

    # If no template, copy header and footer positions
    if not tpath:
        # Copy header region exactly (1..h_dst-1 from src onto dst)
        for r in range(1, h_src):  # up to row above header
            _clone_row(ws_src, r, ws_dst, r)

    # Row style template (from destination first data row if template exists; else from source)
    style_ws = ws_dst if tpath else ws_src
    style_row_idx = data_start_dst if tpath else data_start_src

    # Write roster-ordered rows
    for i, emp_name in enumerate(order):
        r_dst = data_start_dst + i

        # Initialize styles + clear values
        max_cols = max(ws_src.max_column, ws_dst.max_column)
        for c in range(1, max_cols + 1):
            sref = style_ws.cell(row=style_row_idx, column=c)
            d = ws_dst.cell(row=r_dst, column=c)
            _copy_style(sref, d)
            # do not clobber template cell formulas; otherwise blank/zero
            if isinstance(d.value, str) and d.value.startswith("="):
                pass
            else:
                # numeric format heuristic
                d.value = 0 if sref.data_type == 'n' else ""

        # Fill name
        ws_dst.cell(row=r_dst, column=name_dst).value = emp_name
        # Fill SSN from roster CSV (authoritative). If blank, try src row or leave blank.
        if ssn_dst:
            ssn = ssn_for(emp_name, roster_csv)
            if not ssn and emp_name in rows_by_name and ssn_src:
                ssn = ws_src.cell(row=rows_by_name[emp_name], column=ssn_src).value
            ws_dst.cell(row=r_dst, column=ssn_dst).value = ssn

        # Copy source row values if this employee exists in generated WBS
        src_row = rows_by_name.get(emp_name)
        if src_row:
            for c in range(1, max_cols + 1):
                s = ws_src.cell(row=src_row, column=c)
                d = ws_dst.cell(row=r_dst, column=c)
                d.value = s.value
            # Re-enforce name/SSN cells in case the source row differed
            ws_dst.cell(row=r_dst, column=name_dst).value = emp_name
            if ssn_dst:
                ssn = ssn_for(emp_name, roster_csv) or (ws_src.cell(row=src_row, column=ssn_src).value if ssn_src else "")
                ws_dst.cell(row=r_dst, column=ssn_dst).value = ssn
        else:
            # Missing in source: keep numeric columns at 0; let template row formulas persist if present.
            # Optionally fold in Sierra parsed hours if available:
            if not sierra_df.empty:
                row = sierra_df[sierra_df["Name"].str.strip() == emp_name]
                if not row.empty:
                    for col_name, c_idx in (("REGULAR","REGULAR"),("OVERTIME","OVERTIME"),("DOUBLETIME","DOUBLETIME"),("Totals","Totals")):
                        if c_idx in cmap_dst:
                            ws_dst.cell(row=r_dst, column=cmap_dst[c_idx]).value = float(row.iloc[0][col_name])

    # If not using template, copy footer after our last data row
    if not tpath:
        # find first footer row in src
        blanks, r = 0, data_start_src
        last_data_row_src = data_start_src - 1
        while r <= ws_src.max_row and blanks < 12:
            v = ws_src.cell(row=r, column=name_src).value
            if v is None or str(v).strip() == "":
                blanks += 1
            else:
                blanks = 0
                last_data_row_src = r
            r += 1
        footer_start_src = last_data_row_src + 1
        footer_start_dst = data_start_dst + len(order)
        # copy footer rows (values + styles)
        for r_src in range(footer_start_src, ws_src.max_row + 1):
            _clone_row(ws_src, r_src, ws_dst, footer_start_dst + (r_src - footer_start_src))
        # Replace sheet
        if "WEEKLY" in wb_dst.sheetnames:
            del wb_dst["WEEKLY"]
        ws_dst.title = "WEEKLY"
        wb_dst.save(output_wbs_path)
    else:
        # Save the filled template as the final output
        wb_dst.save(output_wbs_path)
