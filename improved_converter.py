# src/improved_converter.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
import logging

logger = logging.getLogger("improved_converter")
logger.setLevel(logging.INFO)

ROOT = Path(__file__).resolve().parent
DATA = ROOT.parent / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

def _norm(s: str) -> str:
    return (s or "").strip().lower()

def _canon_name(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace(".", "")
    s = re.sub(r"\s*,\s*", ", ", s)
    return s.lower()

def _num(v) -> float:
    try:
        if v is None or (isinstance(v, str) and v.strip() == ""):
            return 0.0
        return float(v)
    except Exception:
        return 0.0

def _load_order(path: Path | str = ORDER_TXT) -> List[str]:
    p = Path(path)
    if not p.exists():
        return []
    return [ln.strip() for ln in p.read_text(encoding="utf-8").splitlines() if ln.strip()]

class SierraToWBSConverter:
    def __init__(self, gold_master_order_path: str | None = None):
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = [ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()]

    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """Return DF with Name, REGULAR, OVERTIME, DOUBLETIME, Hours, and __canon."""
        # Try reading likely sheetnames/headers robustly
        tried = []
        def try_read(sheet_name=None, header=7):
            try:
                df = pd.read_excel(input_path, sheet_name=sheet_name, header=header)
                df = df.dropna(how='all')
                return df
            except Exception as e:
                tried.append((sheet_name, header, str(e)))
                return pd.DataFrame()

        df = try_read("WEEKLY", 7)
        if df.empty:
            df = try_read(0, 7)
        if df.empty:
            df = try_read(0, 0)
        if df.empty:
            # last resort: read first sheet without header
            df = pd.read_excel(input_path, sheet_name=0, header=None)
            df = df.dropna(how='all')

        # attempt to locate the "Employee Name" column by fuzzy match
        cols = [str(c) for c in df.columns]
        name_col = None
        for c in cols:
            if "employee" in c.lower() and "name" in c.lower() or c.strip().lower() in ("name", "employee name", "employee"):
                name_col = c; break
        if not name_col:
            # fallback to second or third column if obvious
            if len(cols) >= 2:
                name_col = cols[1]
            else:
                name_col = cols[0]

        # attempt to find REGULAR, OVERTIME, DOUBLETIME by aliases
        def find_column_alias(dfcols, aliases):
            for a in aliases:
                for c in dfcols:
                    if str(c).strip().lower() == a.lower() or a.lower() in str(c).strip().lower():
                        return c
            return None

        reg_col = find_column_alias(cols, ["REGULAR", "A01", "Regular"])
        ot_col  = find_column_alias(cols, ["OVERTIME", "A02", "Overtime", "OT"])
        dt_col  = find_column_alias(cols, ["DOUBLETIME", "A03", "Double Time"])

        # build out DataFrame
        out = pd.DataFrame()
        out["Name"] = df[name_col].astype(str).str.strip() if name_col in df.columns else df.iloc[:, 0].astype(str).str.strip()

        def to_num_series(col):
            if col is None:
                return pd.Series([0.0] * len(out))
            return pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        out["REGULAR"] = to_num_series(reg_col)
        out["OVERTIME"] = to_num_series(ot_col)
        out["DOUBLETIME"] = to_num_series(dt_col)
        out["Hours"] = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out = out[out["Name"].astype(str).str.strip().str.len() > 0].reset_index(drop=True)
        out["__canon"] = out["Name"].map(_canon_name)
        return out

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        p = Path(ROSTER_CSV)
        if not p.exists():
            logger.warning("Roster not found: %s", p)
            return {}
        try:
            df = pd.read_csv(p)
        except Exception as e:
            logger.warning("Failed to read roster: %s", e)
            return {}
        # alias mapping
        cols = {c: c.lower().replace(" ", "") for c in df.columns}
        def find(*aliases):
            for a in aliases:
                for raw, norm in cols.items():
                    if norm == a.lower().replace(" ", ""):
                        return raw
            return None
        name_col = find("employeename", "name")
        ssn_col = find("ssn", "socialsecurity")
        type_col = find("type")
        dept_col = find("dept", "department")
        pay_col = find("payrate", "rate")
        out = {}
        if name_col is None:
            return out
        for _, r in df.iterrows():
            nm = str(r.get(name_col, "")).strip()
            if not nm:
                continue
            key = _canon_name(nm)
            out[key] = {
                "ssn": str(r.get(ssn_col, "")).strip() if ssn_col else "",
                "type": str(r.get(type_col, "")).strip() if type_col else "",
                "dept": str(r.get(dept_col, "")).strip() if dept_col else "",
                "pay_rate": str(r.get(pay_col, "")).strip() if pay_col else ""
            }
        return out

    def convert(self, input_path: str, output_path: str) -> Dict:
        try:
            order = self.gold_master_order[:] or _load_order()
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}
            df = self.parse_sierra_file(input_path)
            sierra_map = {row["__canon"]: row for _, row in df.iterrows()}
            roster = self._load_roster()

            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            wb = load_workbook(TEMPLATE_XLSX)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws = wb[TARGET_SHEET]

            # locate header row and column positions
            header_row = None
            cmap = {}
            for r in range(1, min(ws.max_row, 150) + 1):
                labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
                low = [lbl.replace(" ", "").lower() for lbl in labels]
                if any("employeename" == x or "employee" == x for x in low):
                    header_row = r
                    for i, lbl in enumerate(labels, start=1):
                        k0 = lbl.replace(" ", "").lower()
                        if "employeename" in k0 or "employee" == k0:
                            cmap["Employee Name"] = i
                        elif k0 in ("ssn", "socialsecuritynumber", "socialsecurity#"):
                            cmap["SSN"] = i
                        elif "regular" in k0 or k0 == "a01":
                            cmap["REGULAR"] = i
                        elif "overtime" in k0 or k0 == "a02" or "ot" in k0:
                            cmap["OVERTIME"] = i
                        elif "double" in k0 or k0 == "a03":
                            cmap["DOUBLETIME"] = i
                        elif "status" in k0:
                            cmap["Status"] = i
                        elif "type" in k0:
                            cmap["Type"] = i
                        elif "payrate" in k0 or "pay" in k0:
                            cmap["Pay Rate"] = i
                        elif "dept" in k0 or "department" in k0:
                            cmap["Dept"] = i
                        elif "total" in k0 or "totals" in k0:
                            cmap["Totals"] = i
                    break

            if header_row is None:
                return {"success": False, "error": "could not find header row in template"}

            start = header_row + 1
            name_col = cmap.get("Employee Name")
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
            otL  = get_column_letter(ot_col) if ot_col else None
            dtL  = get_column_letter(dt_col) if dt_col else None

            matches = 0
            for i, emp in enumerate(order):
                r = start + i
                ws.cell(row=r, column=name_col).value = emp
                k = _canon_name(emp)
                s = sierra_map.get(k)
                ro = roster.get(k, {})

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

                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0
                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt

                if s is not None:
                    matches += 1

                if tot_col:
                    c = ws.cell(row=r, column=tot_col)
                    has_formula = isinstance(c.value, str) and c.value.startswith("=")
                    if not has_formula and regL and otL and dtL:
                        c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

            wb.save(output_path)

            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return {"success": True, "employees": len(order), "hours": total_hours, "matched_rows": matches}
        except Exception as e:
            return {"success": False, "error": str(e)}
