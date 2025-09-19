# -*- coding: utf-8 -*-
"""
Improved Sierra → WBS converter.
- Reads Sierra "WEEKLY" sheet (or first sheet fallback)
- Merges gold roster (SSN/Status/Type/Dept/Pay Rate)
- Writes into data/wbs_template.xlsx while preserving layout
- Ensures far-right Totals column is populated and a grand totals row exists
"""

from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Any
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
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

# ---------- helpers ----------
def _norm(s: Any) -> str:
    return ("" if s is None else str(s)).strip().lower()

def _canon_name(s: Any) -> str:
    s = "" if s is None else str(s)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace(".", "")
    s = re.sub(r"\s*,\s*", ", ", s)
    s = s.replace(" ,", ",")
    return s.lower()

def _num(v: Any) -> float:
    try:
        if v is None:
            return 0.0
        if isinstance(v, str) and v.strip() == "":
            return 0.0
        return float(v)
    except Exception:
        return 0.0

def _load_order() -> List[str]:
    if not ORDER_TXT.exists():
        return []
    return [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """Find the row that contains 'Employee Name' and map important columns."""
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 150) + 1):
        labels = ["" if c.value is None else str(c.value).strip() for c in ws[r]]
        low = [_norm(v).replace(" ", "") for v in labels]
        if "employeename" in low:
            for i, v in enumerate(labels, start=1):
                k = _norm(v); k0 = k.replace(" ", "")
                if k0 == "employeename":
                    cmap["Employee Name"] = i
                elif k0 in ("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno"):
                    cmap["SSN"] = i
                elif k0 == "regular":
                    cmap["REGULAR"] = i
                elif k0 in ("overtime", "ot"):
                    cmap["OVERTIME"] = i
                elif k0 in ("doubletime", "doubletime", "double time"):
                    cmap["DOUBLETIME"] = i
                elif k0 == "status":
                    cmap["Status"] = i
                elif k0 == "type":
                    cmap["Type"] = i
                elif k0 in ("payrate", "pay", "payrate", "pay rate"):
                    cmap["Pay Rate"] = i
                elif k0 in ("dept", "department"):
                    cmap["Dept"] = i
                elif k0 in ("totals", "total", "sum"):
                    cmap["Totals"] = i
            if "Employee Name" not in cmap:
                continue
            return r, cmap
    raise ValueError("Could not locate header row containing 'Employee Name'")

def _first_data_row(h: int) -> int:
    return h + 1

# ---------- converter ----------
class SierraToWBSConverter:
    """
    Fills Name, SSN, REG/OT/DT, and (from roster) Dept/Type/Pay Rate into data/wbs_template.xlsx.
    Retains template formatting and formulas. Exposes .gold_master_order for health logs.
    """

    def __init__(self, gold_master_order_path: str | None = None):
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = [
                ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()
            ]

    # ---------- used by validation ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """Return DF with Name, REGULAR, OVERTIME, DOUBLETIME, Hours, and __canon."""
        # Try Sierra export layout first (header at row 8)
        try:
            df = pd.read_excel(input_path, sheet_name="WEEKLY", header=7)
        except Exception:
            df = pd.read_excel(input_path, sheet_name=0, header=7)
        df = df.dropna(how="all")

        # Sometimes row 1 repeats header text again
        if not df.empty and "Employee Name" in df.columns and str(df.iloc[0]["Employee Name"]).strip() == "Employee Name":
            df = df.iloc[1:]

        # If columns still don’t make sense, fall back to row 1 header
        expected_any = {"Employee Name", "Name", "REGULAR", "OVERTIME", "DOUBLETIME"}
        if len(expected_any.intersection(set(map(str, df.columns)))) < 2:
            try:
                df = pd.read_excel(input_path, sheet_name=0, header=0).dropna(how="all")
            except Exception:
                pass

        name_col = None
        for cand in ["Employee Name", "Name", "Unnamed: 2"]:
            if cand in df.columns:
                name_col = cand
                break
        if not name_col:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})

        def to_num_series(series_like):
            return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

        src_cols = {
            "REGULAR": ["REGULAR", "Regular"],
            "OVERTIME": ["OVERTIME", "Overtime", "OT"],
            "DOUBLETIME": ["DOUBLETIME", "Double Time", "DOUBLE TIME"],
        }
        for dst, src_list in src_cols.items():
            col = next((c for c in src_list if c in df.columns), None)
            out[dst] = to_num_series(df[col]) if col else 0.0

        out["Hours"] = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out = out[out["Name"].astype(str).str.strip().str.len() > 0]
        out["__canon"] = out["Name"].map(_canon_name)
        return out

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        """Return canonical_name -> {ssn,status,type,dept,pay_rate}."""
        if not ROSTER_CSV.exists():
            print(f"[WARN] Roster not found: {ROSTER_CSV}")
            return {}
        try:
            df = pd.read_csv(ROSTER_CSV)
        except Exception as e:
            print(f"[WARN] Failed to read roster CSV: {e}")
            return {}

        header_map = {c: _norm(c).replace(" ", "") for c in df.columns}

        def _pick(*aliases):
            aliases = [a.replace(" ", "").lower() for a in aliases]
            for c, norm in header_map.items():
                if norm in aliases:
                    return c
            return None

        name_col = _pick("employeename", "name")
        ssn_col  = _pick("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno")
        stat_col = _pick("status")
        type_col = _pick("type")
        dept_col = _pick("dept", "department")
        rate_col = _pick("payrate", "payrate", "rate", "pay")

        if not name_col:
            print("[WARN] Roster missing 'Employee Name' column (alias).")
            return {}

        out: Dict[str, Dict[str, str]] = {}
        for _, r in df.iterrows():
            nm = str(r.get(name_col, "")).strip()
            if not nm:
                continue
            k = _canon_name(nm)
            out[k] = {
                "ssn": "" if ssn_col is None or pd.isna(r.get(ssn_col)) else str(r.get(ssn_col)).strip(),
                "status": "" if stat_col is None or pd.isna(r.get(stat_col)) else str(r.get(stat_col)).strip(),
                "type": "" if type_col is None or pd.isna(r.get(type_col)) else str(r.get(type_col)).strip(),
                "dept": "" if dept_col is None or pd.isna(r.get(dept_col)) else str(r.get(dept_col)).strip(),
                "pay_rate": "" if rate_col is None or pd.isna(r.get(rate_col)) else str(r.get(rate_col)).strip(),
            }
        if not out:
            print("[WARN] Roster parsed but produced 0 rows after normalization.")
        return out

    # ---------- main convert ----------
    def convert(self, input_path: str, output_path: str) -> Dict[str, Any]:
        try:
            # gold order
            order = self.gold_master_order[:] or _load_order()
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}

            # Sierra values
            df = self.parse_sierra_file(input_path)
            sierra_map = {row["__canon"]: row for _, row in df.iterrows()}

            # roster attributes
            roster = self._load_roster()

            # open template
            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            # locate columns in template
            h, cmap = _find_header_row(ws)
            start = _first_data_row(h)
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

            matches = 0
            for i, emp in enumerate(order):
                r = start + i
                ws.cell(row=r, column=name_col).value = emp

                k = _canon_name(emp)
                s = sierra_map.get(k)
                ro = roster.get(k, {})

                # From roster
                if ssn_col:  ws.cell(row=r, column=ssn_col).value  = ro.get("ssn", "")
                if stat_col: ws.cell(row=r, column=stat_col).value = ro.get("status", "") or "A"
                if type_col: ws.cell(row=r, column=type_col).value = ro.get("type", "") or "H"
                if dept_col: ws.cell(row=r, column=dept_col).value = ro.get("dept", "")

                if rate_col:
                    try:
                        val = float(ro.get("pay_rate", "") or 0.0)
                    except Exception:
                        val = 0.0
                    ws.cell(row=r, column=rate_col).value = val

                # Hours from Sierra
                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0
                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt

                if s is not None:
                    matches += 1

                # Ensure per-row Totals column has a formula if blank
                if tot_col and regL and otL and dtL:
                    c = ws.cell(row=r, column=tot_col)
                    if not (isinstance(c.value, str) and c.value.startswith("=")):
                        c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

            # Append a totals row at the first empty row after last employee
            last_row = start + len(order)
            if reg_col and ot_col and dt_col and tot_col:
                ws.cell(row=last_row, column=name_col).value = "TOTAL"
                # Sum each hours column (keep template formats)
                for col in (reg_col, ot_col, dt_col, tot_col):
                    colL = get_column_letter(col)
                    ws.cell(row=last_row, column=col).value = f"=SUM({colL}{start}:{colL}{last_row-1})"

            wb.save(output_path)

            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            if matches == 0:
                print("[WARN] No Sierra rows matched names in gold order after canonicalization.")
            if not roster:
                print("[WARN] Roster merge skipped (missing or unreadable).")

            return {"success": True, "employees": len(order), "hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
