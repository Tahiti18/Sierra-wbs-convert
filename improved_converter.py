# improved_converter.py
from __future__ import annotations
import os
from pathlib import Path
from typing import Dict, List, Tuple
from copy import copy

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from openpyxl.cell.cell import Cell


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"       # one name per line
ROSTER_CSV = DATA / "gold_master_roster.csv"     # must include "Employee Name" and "SSN"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"       # gold layout
TARGET_SHEET = "WEEKLY"


# ------------------------ small utils ------------------------

def _norm_header(h: str) -> str:
    return (h or "").strip().lower().replace(" ", "")

def _load_order() -> List[str]:
    if not ORDER_TXT.exists():
        return []
    names = [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines()]
    return [n for n in names if n]

def _load_roster_ssn() -> Dict[str, str]:
    if not ROSTER_CSV.exists():
        return {}
    df = pd.read_csv(ROSTER_CSV)
    # normalize headers
    cols = {c: _norm_header(c) for c in df.columns}
    name_col = next((c for c, n in cols.items() if n in ("employeename", "name")), None)
    ssn_col  = next((c for c, n in cols.items() if n == "ssn"), None)
    if not name_col or not ssn_col:
        return {}
    out = {}
    for _, r in df[[name_col, ssn_col]].iterrows():
        nm = str(r[name_col]).strip()
        sv = "" if pd.isna(r[ssn_col]) else str(r[ssn_col]).strip()
        if nm:
            out[nm] = sv
    return out

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    colmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 80) + 1):
        vals = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [_norm_header(v) for v in vals]
        if any(x in ("employeename",) for x in low):
            for idx, v in enumerate(vals, start=1):
                k = _norm_header(v)
                if k in ("employeename", "employee name"): colmap["Employee Name"] = idx
                elif k == "ssn": colmap["SSN"] = idx
                elif k == "regular": colmap["REGULAR"] = idx
                elif k == "overtime": colmap["OVERTIME"] = idx
                elif k == "doubletime": colmap["DOUBLETIME"] = idx
                elif k == "totals": colmap["Totals"] = idx
            return r, colmap
    raise ValueError("Could not find header row with 'Employee Name'")

def _first_data_row(header_row: int) -> int:
    return header_row + 1

def _collect_rows_by_name(ws: Worksheet, name_col: int, start_row: int) -> Dict[str, int]:
    out: Dict[str, int] = {}
    blanks = 0
    r = start_row
    while r <= ws.max_row and blanks < 12:
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

def _ensure_rows(ws: Worksheet, data_start: int, needed: int, style_row: int, name_col: int) -> int:
    """ensure destination sheet has >= needed rows in the data region; returns footer_start row"""
    # find current footer start
    r = data_start
    blanks = 0
    last = data_start - 1
    while r <= ws.max_row and blanks < 10:
        val = ws.cell(row=r, column=name_col).value
        if val is None or str(val).strip() == "":
            blanks += 1
        else:
            blanks = 0
            last = r
        r += 1
    footer_start = last + 1
    available = max(0, footer_start - data_start)
    if available >= needed:
        return footer_start
    to_insert = needed - available
    ws.insert_rows(footer_start, amount=to_insert)
    for i in range(to_insert):
        row_idx = footer_start + i
        for c in range(1, ws.max_column + 1):
            s = ws.cell(row=style_row, column=c)
            d = ws.cell(row=row_idx, column=c)
            d.value = ""
            _copy_style(s, d)
    return footer_start + to_insert

def _num(val) -> float:
    try:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return 0.0
        return float(val)
    except Exception:
        return 0.0


# ------------------------ core class ------------------------

class SierraToWBSConverter:
    """
    Converter that:
      - parses Sierra input for weekly values,
      - writes values into data/wbs_template.xlsx,
      - enforces gold order and SSNs.
    """

    def __init__(self, _gold_master_path_unused: str | None = None):
        # paths are fixed to /data; _gold_master_path_unused kept for compatibility
        pass

    # --------- public: quick parse for /validate ---------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Returns a DF with at least columns: Name, REGULAR, OVERTIME, DOUBLETIME, Hours (Hours = sum(REG+OT+DT))
        """
        # Sierra file has a sheet named WEEKLY with headers around row 8
        df = pd.read_excel(input_path, sheet_name="WEEKLY", header=7)
        df = df.dropna(how="all")
        # some files repeat header; drop if found
        if (df.iloc[0:1].astype(str).apply(lambda x: (x == 'Employee Name').any(), axis=1).any()):
            df = df.iloc[1:]

        # detect name column
        name_col = None
        for cand in ["Employee Name", "Unnamed: 2", "Name"]:
            if cand in df.columns:
                name_col = cand
                break
        if not name_col:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})
        for col in ["REGULAR", "OVERTIME", "DOUBLETIME"]:
            out[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0.0)
        out["Hours"] = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out = out[out["Name"].str.len() > 0]
        return out

    # --------- public: main convert ----------
    def convert(self, input_path: str, output_path: str) -> Dict:
        """
        Writes a final WBS workbook at output_path using the gold template and roster.
        """
        try:
            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}

            order = _load_order()
            if not order:
                return {"success": False, "error": f"gold_master_order.txt is empty or missing"}

            ssn_map = _load_roster_ssn()  # may be partial; blanks allowed

            # Parse Sierra (only for values; not for order or SSNs)
            df = self.parse_sierra_file(input_path)

            # Open template
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            # Locate headers/columns
            h_row, cmap = _find_header_row(ws)
            data_start = _first_data_row(h_row)
            name_col = cmap["Employee Name"]
            ssn_col = cmap.get("SSN")
            reg_col = cmap.get("REGULAR")
            ot_col  = cmap.get("OVERTIME")
            dt_col  = cmap.get("DOUBLETIME")
            tot_col = cmap.get("Totals")

            # Ensure enough rows
            footer_start = _ensure_rows(ws, data_start, len(order), data_start, name_col)

            # Build quick lookup from Sierra by name
            df_norm = df.copy()
            df_norm["Name"] = df_norm["Name"].astype(str).str.strip()
            sierra_map = {r["Name"]: r for _, r in df_norm.iterrows()}

            # Fill rows in exact order
            for i, emp in enumerate(order):
                r = data_start + i

                # clear row values first, keeping formulas
                for c in range(1, ws.max_column + 1):
                    cell = ws.cell(row=r, column=c)
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        continue
                    # zero numerics, blank others
                    if cell.data_type == 'n':
                        cell.value = 0
                    else:
                        cell.value = ""

                # enforce name + SSN
                ws.cell(row=r, column=name_col).value = emp
                if ssn_col is not None:
                    ws.cell(row=r, column=ssn_col).value = ssn_map.get(emp, "")

                # values from Sierra (if present)
                srow = sierra_map.get(emp)
                if srow is not None:
                    if reg_col:
                        ws.cell(row=r, column=reg_col).value = _num(srow.get("REGULAR", 0))
                    if ot_col:
                        ws.cell(row=r, column=ot_col).value = _num(srow.get("OVERTIME", 0))
                    if dt_col:
                        ws.cell(row=r, column=dt_col).value = _num(srow.get("DOUBLETIME", 0))
                    # Totals column: write ONLY if the template cell is NOT a formula
                    if tot_col:
                        tot_cell = ws.cell(row=r, column=tot_col)
                        if not (isinstance(tot_cell.value, str) and str(tot_cell.value).startswith("=")):
                            ws.cell(row=r, column=tot_col).value = _num(srow.get("REGULAR", 0)) + _num(srow.get("OVERTIME", 0)) + _num(srow.get("DOUBLETIME", 0))

            # Save as output
            wb.save(output_path)

            # quick stats for API
            total_hours = float(df_norm[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df_norm.empty else 0.0
            return {"success": True, "employees": len(order), "hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
