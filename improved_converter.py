#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Improved Sierra → WBS converter
- Robust parser for Sierra raw timesheets and WBS-like sheets
- Uses data/gold_master_order.txt for row order
- Optionally uses data/gold_master_roster.csv for SSN/Status/Type/Dept/Pay Rate
- Fills an existing template at data/wbs_template.xlsx and preserves formatting
"""

from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# ---------- paths ----------
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"       # optional
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"         # recommended
TARGET_SHEET = "WEEKLY"                            # in template

# ---------- helpers ----------
def _norm(s: str) -> str:
    return (s or "").strip().lower()

def _canon_name(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace(".", "")
    s = re.sub(r"\s*,\s*", ", ", s)
    s = s.replace(" ,", ",")
    return s.lower()

def _num(v) -> float:
    try:
        if v is None:
            return 0.0
        if isinstance(v, str) and v.strip() == "":
            return 0.0
        return float(v)
    except Exception:
        return 0.0

def _load_order() -> List[str]:
    if ORDER_TXT.exists():
        return [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines() if ln.strip()]
    return []

def _pick_column(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    cols = {c: _norm(c).replace(" ", "") for c in df.columns}
    targets = [a.replace(" ", "").lower() for a in aliases]
    for c, n in cols.items():
        if n in targets:
            return c
    return None

def _safe_read_xlsx(path: str, sheet, header) -> Optional[pd.DataFrame]:
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header)
        if isinstance(df, pd.DataFrame):
            return df.dropna(how="all")
    except Exception:
        return None
    return None

def _best_sheet(path: str) -> Tuple[pd.DataFrame, int]:
    """
    Return (df, header_row_used). We try multiple sheet/header combos and pick the
    frame that actually contains a name column.
    """
    tries = [
        ("WEEKLY", 7), ("WEEKLY", 8),
        (0, 7), (0, 8),
        (0, 0), ("WEEKLY", 0),
    ]
    candidates: List[Tuple[pd.DataFrame, int]] = []
    for sheet, header in tries:
        df = _safe_read_xlsx(path, sheet, header)
        if df is None or df.empty:
            continue
        name_col = _pick_column(df, ["employee name", "name"])
        if name_col:
            candidates.append((df, header))
    if candidates:
        # choose the widest (most columns) that contains any hours columns
        def score(item):
            df, _ = item
            hits = 0
            for a in (["regular", "a01"], ["overtime", "ot", "a02"], ["doubletime", "double time", "a03"], ["hours"]):
                if _pick_column(df, a):
                    hits += 1
            return (hits, df.shape[1], df.shape[0])
        candidates.sort(key=score, reverse=True)
        return candidates[0]
    # fallback first sheet header 0
    df = _safe_read_xlsx(path, 0, 0)
    if df is None:
        return pd.DataFrame(), 0
    return df, 0

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Find header row in the *template* where "Employee Name" exists.
    Build a column map for SSN, Employee Name, REGULAR, OVERTIME, DOUBLETIME, etc.
    """
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 150) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        lows = [_norm(v).replace(" ", "") for v in labels]
        if "employeename" in lows:
            for i, v in enumerate(labels, start=1):
                k = _norm(v).replace(" ", "")
                if k == "employeename":
                    cmap["Employee Name"] = i
                elif k in ("ssn", "socialsecuritynumber", "socialsecurity#"):
                    cmap["SSN"] = i
                elif k == "status":
                    cmap["Status"] = i
                elif k == "type":
                    cmap["Type"] = i
                elif k in ("payrate", "pay", "payrate"):
                    cmap["Pay Rate"] = i
                elif k in ("dept", "department"):
                    cmap["Dept"] = i
                elif k == "regular" or k == "a01":
                    cmap["REGULAR"] = i
                elif k in ("overtime", "ot", "a02"):
                    cmap["OVERTIME"] = i
                elif k in ("doubletime", "doubletime", "a03"):
                    cmap["DOUBLETIME"] = i
                elif k in ("totals", "total", "sum"):
                    cmap["Totals"] = i
            return r, cmap
    raise ValueError("Template header row not found (missing 'Employee Name').")

def _first_data_row(header_row: int) -> int:
    return header_row + 1

# ---------- converter ----------
class SierraToWBSConverter:
    """
    Public API used by main.py
    - .gold_master_order is kept for /api/health
    - parse_sierra_file(): used by /api/validate-sierra-file
    - convert(): used by /api/process-payroll
    """
    def __init__(self, gold_master_order_path: Optional[str] = None):
        # expose for /health
        self.gold_master_order: List[str] = []
        src = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if src.exists():
            self.gold_master_order = [
                ln.strip()
                for ln in src.read_text(encoding="utf-8").splitlines()
                if ln.strip()
            ]

    # ---------- VALIDATION ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Normalize uploaded Excel into:
        columns = [Name, REGULAR, OVERTIME, DOUBLETIME, Hours, __canon]
        Works with Sierra raw timesheets OR WBS-like layouts.
        """
        df, _hdr = _best_sheet(input_path)
        if df is None or df.empty:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        # identify columns by aliases
        name_col = _pick_column(df, ["employee name", "name"])
        reg_col  = _pick_column(df, ["regular", "a01"])
        ot_col   = _pick_column(df, ["overtime", "ot", "a02"])
        dt_col   = _pick_column(df, ["doubletime", "double time", "a03"])
        hrs_col  = _pick_column(df, ["hours"])  # sometimes present in Sierra raw rows

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()}) if name_col else pd.DataFrame({"Name": []})

        def to_num(series_like):
            return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

        # If we have A01/A02/A03 or REG/OT/DT use them; otherwise derive from "Hours"
        out["REGULAR"]   = to_num(df[reg_col]) if reg_col else 0.0
        out["OVERTIME"]  = to_num(df[ot_col])  if ot_col  else 0.0
        out["DOUBLETIME"]= to_num(df[dt_col])  if dt_col  else 0.0

        # If all three are zero and there is a plain Hours column, use it as REGULAR
        if (out[["REGULAR","OVERTIME","DOUBLETIME"]].sum(axis=1) == 0).all() and hrs_col:
            out["REGULAR"] = to_num(df[hrs_col])

        # strip empties & headers that sometimes repeat inside the sheet
        out = out[out["Name"].astype(str).str.strip().ne("")]
        out = out[~out["Name"].str.lower().str.contains(r"^employee\s*name$|^name$", regex=True)]
        out["__canon"] = out["Name"].map(_canon_name)
        out["Hours"] = out[["REGULAR","OVERTIME","DOUBLETIME"]].sum(axis=1)

        # If there are obviously too many rows (e.g., 500+ due to template noise),
        # keep only names that appear in the gold order to avoid bogus counts.
        order = set(n.lower() for n in (self.gold_master_order or _load_order()))
        if len(out) > 200 and order:
            out = out[out["__canon"].isin({_canon_name(n) for n in order})]

        # Final clean
        out = out[out["Hours"] > 0].reset_index(drop=True)
        return out

    # ---------- ROSTER ----------
    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        """Name → {ssn,status,type,dept,pay_rate}, using permissive header aliases. Optional."""
        if not ROSTER_CSV.exists():
            print(f"[WARN] Roster not found: {ROSTER_CSV}")
            return {}
        try:
            df = pd.read_csv(ROSTER_CSV)
        except Exception as e:
            print(f"[WARN] Failed to read roster CSV: {e}")
            return {}

        def pick(*aliases):
            return _pick_column(df, list(aliases))

        name_col = pick("employee name", "name")
        ssn_col  = pick("ssn", "social security number", "socialsecurity#", "socialsecurityno")
        stat_col = pick("status")
        type_col = pick("type")
        dept_col = pick("dept", "department")
        rate_col = pick("pay rate", "payrate", "rate")

        if not name_col:
            print("[WARN] Roster missing 'Employee Name' column.")
            return {}

        out: Dict[str, Dict[str, str]] = {}
        for _, r in df.iterrows():
            nm = str(r.get(name_col, "")).strip()
            if not nm:
                continue
            k = _canon_name(nm)
            def val(c):
                v = None if c is None else r.get(c)
                if pd.isna(v):
                    return ""
                return str(v).strip()
            out[k] = {
                "ssn": val(ssn_col),
                "status": val(stat_col) or "A",
                "type": val(type_col) or "H",
                "dept": val(dept_col),
                "pay_rate": val(rate_col),
            }
        return out

    # ---------- PROCESS ----------
    def convert(self, input_path: str, output_path: str) -> Dict:
        """
        Fill data/wbs_template.xlsx keeping formatting.
        """
        try:
            order = self.gold_master_order[:] or _load_order()
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}

            df = self.parse_sierra_file(input_path)
            # map Sierra by canonical name
            sierra_map = {row["__canon"]: row for _, row in df.iterrows()}

            roster = self._load_roster()  # optional; never fails convert()

            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}

            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            header_row, cmap = _find_header_row(ws)
            start = _first_data_row(header_row)

            name_col = cmap["Employee Name"]
            ssn_col  = cmap.get("SSN")
            reg_col  = cmap.get("REGULAR")
            ot_col   = cmap.get("OVERTIME")
            dt_col   = cmap.get("DOUBLETIME")
            rate_col = cmap.get("Pay Rate")
            stat_col = cmap.get("Status")
            type_col = cmap.get("Type")
            dept_col = cmap.get("Dept")
            tot_col  = cmap.get("Totals")

            regL = get_column_letter(reg_col) if reg_col else None
            otL  = get_column_letter(ot_col)  if ot_col  else None
            dtL  = get_column_letter(dt_col)  if dt_col  else None

            # clear existing body region (up to number of names in order)
            for i in range(len(order)):
                r = start + i
                for c in range(1, ws.max_column + 1):
                    # don't wipe formatting; set to None only for known data columns
                    pass  # we will overwrite target cells directly below

            # write rows strictly in gold order
            matched = 0
            for i, emp in enumerate(order):
                r = start + i
                ws.cell(row=r, column=name_col).value = emp

                key = _canon_name(emp)
                s = sierra_map.get(key)
                ro = roster.get(key, {})

                if ssn_col:  ws.cell(row=r, column=ssn_col).value  = ro.get("ssn", "")
                if stat_col: ws.cell(row=r, column=stat_col).value = ro.get("status", "") or "A"
                if type_col: ws.cell(row=r, column=type_col).value = ro.get("type", "") or "H"
                if dept_col: ws.cell(row=r, column=dept_col).value = ro.get("dept", "")

                if rate_col:
                    try:
                        rate_val = float(ro.get("pay_rate", "") or 0.0)
                    except Exception:
                        rate_val = 0.0
                    ws.cell(row=r, column=rate_col).value = rate_val

                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0

                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt

                if s is not None:
                    matched += 1

                # Totals (pink): if template cell lacks a formula, supply =REG+OT+DT
                if tot_col and regL and otL and dtL:
                    c = ws.cell(row=r, column=tot_col)
                    cv = c.value
                    if not (isinstance(cv, str) and cv.startswith("=")):
                        c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

            wb.save(output_path)

            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return {
                "success": True,
                "employees_processed": len(order),
                "total_hours": total_hours,
                "matched": matched
            }

        except Exception as e:
            return {"success": False, "error": str(e)}
