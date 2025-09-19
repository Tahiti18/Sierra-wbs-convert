# improved_converter.py — fills hours/rates, robust name match, totals formula
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

# --------------------- helpers ---------------------
def _norm(s: str) -> str:
    return (s or "").strip().lower()

def _canon_name(s: str) -> str:
    """
    Canonicalize 'Employee Name' so Sierra vs Gold order map reliably.
    - trim spaces, collapse multiple spaces
    - unify case
    - remove periods, extra commas/spaces around commas
    - remove double spaces and stray middle-initial punctuation
    """
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)             # collapse spaces
    s = s.replace(".", "")
    s = re.sub(r"\s*,\s*", ", ", s)        # clean comma spacing
    s = s.replace(" ,", ",")
    return s.lower()

def _num(v) -> float:
    try:
        if v is None or (isinstance(v, str) and v.strip() == ""):
            return 0.0
        return float(v)
    except Exception:
        return 0.0

def _load_order() -> List[str]:
    if not ORDER_TXT.exists():
        return []
    return [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 120) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [_norm(v).replace(" ", "") for v in labels]
        if "employeename" in low:
            for i, v in enumerate(labels, start=1):
                k = _norm(v)
                k0 = k.replace(" ", "")
                if k0 in ("employeename",):
                    cmap["Employee Name"] = i
                elif k0 in ("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno"):
                    cmap["SSN"] = i
                elif k0 == "regular":
                    cmap["REGULAR"] = i
                elif k0 in ("overtime", "ot"):
                    cmap["OVERTIME"] = i
                elif k0 in ("doubletime", "doubletime", "doubletime", "doubletime", "doubletime", "doubletime", "doubletime", "doubletime", "doubletime"):  # guard
                    cmap["DOUBLETIME"] = i
                elif k0 in ("status",):
                    cmap["Status"] = i
                elif k0 in ("type",):
                    cmap["Type"] = i
                elif k0 in ("payrate", "pay rate"):
                    cmap["Pay Rate"] = i
                elif k0 in ("dept", "department"):
                    cmap["Dept"] = i
                elif k0 in ("totals", "total", "sum"):
                    cmap["Totals"] = i
            return r, cmap
    raise ValueError("Could not locate header row containing 'Employee Name'")

def _first_data_row(h: int) -> int:
    return h + 1

# --------------------- converter ---------------------
class SierraToWBSConverter:
    """
    Opens data/wbs_template.xlsx and fills Name, SSN, REG/OT/DT, Pay Rate (if roster has it)
    using gold order + roster. Keeps headers/format/totals. Compatible with current main.py
    (exposes .gold_master_order).
    """
    def __init__(self, gold_master_order_path: str | None = None):
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = [
                ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()
            ]

    # used by /validate
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        # Try 'WEEKLY'; fallback to first sheet; fallback headers row 0
        try:
            df = pd.read_excel(input_path, sheet_name="WEEKLY", header=7)
        except Exception:
            df = pd.read_excel(input_path, sheet_name=0, header=7)
        df = df.dropna(how="all")

        # guard duplicated header row
        if not df.empty and "Employee Name" in df.columns and str(df.iloc[0]["Employee Name"]).strip() == "Employee Name":
            df = df.iloc[1:]

        # If we still don't see columns, retry with header=0
        expected_any = set(["Employee Name", "Name", "REGULAR", "OVERTIME", "DOUBLETIME"])
        if len(expected_any.intersection(set(map(str, df.columns)))) < 2:
            try:
                df = pd.read_excel(input_path, sheet_name=0, header=0).dropna(how="all")
            except Exception:
                pass

        # Name column
        name_col = None
        for cand in ["Employee Name", "Name", "Unnamed: 2"]:
            if cand in df.columns:
                name_col = cand; break
        if not name_col:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})
        # map common variants just in case
        for src, dst in [("REGULAR", "REGULAR"), ("Overtime", "OVERTIME"), ("OVERTIME", "OVERTIME"),
                         ("Double Time", "DOUBLETIME"), ("DOUBLETIME", "DOUBLETIME")]:
            if src in df.columns and dst not in out.columns:
                out[dst] = pd.to_numeric(df[src], errors="coerce")
        for col in ["REGULAR", "OVERTIME", "DOUBLETIME"]:
            if col not in out.columns:
                out[col] = pd.to_numeric(df.get(col, 0), errors="coerce")
            out[col] = out[col].fillna(0.0)

        out["Hours"] = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out = out[out["Name"].astype(str).str.strip().str.len() > 0]

        # add canonical name for robust matching
        out["__canon"] = out["Name"].map(_canon_name)
        return out

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        """
        Return name -> {ssn, status, type, dept, pay_rate} using canonical name as key.
        """
        if not ROSTER_CSV.exists():
            return {}
        df = pd.read_csv(ROSTER_CSV)
        # try to find columns by loose names
        cols_map = {c: _norm(c) for c in df.columns}
        name_col = next((c for c, n in cols_map.items() if n in ("employeename", "name")), None)
        ssn_col = next((c for c, n in cols_map.items() if n == "ssn"), None)
        stat_col = next((c for c, n in cols_map.items() if n == "status"), None)
        type_col = next((c for c, n in cols_map.items() if n == "type"), None)
        dept_col = next((c for c, n in cols_map.items() if n in ("dept", "department")), None)
        rate_col = next((c for c, n in cols_map.items() if n in ("payrate", "pay rate", "rate")), None)

        out: Dict[str, Dict[str, str]] = {}
        for _, r in df.iterrows():
            nm = str(r.get(name_col, "")).strip()
            if not nm:
                continue
            k = _canon_name(nm)
            out[k] = {
                "ssn": "" if pd.isna(r.get(ssn_col)) else str(r.get(ssn_col)).strip(),
                "status": "" if stat_col is None else ("" if pd.isna(r.get(stat_col)) else str(r.get(stat_col)).strip()),
                "type": "" if type_col is None else ("" if pd.isna(r.get(type_col)) else str(r.get(type_col)).strip()),
                "dept": "" if dept_col is None else ("" if pd.isna(r.get(dept_col)) else str(r.get(dept_col)).strip()),
                "pay_rate": "" if rate_col is None else ("" if pd.isna(r.get(rate_col)) else str(r.get(rate_col)).strip()),
            }
        return out

    def convert(self, input_path: str, output_path: str) -> Dict:
        try:
            order = self.gold_master_order[:] or _load_order()
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}

            df = self.parse_sierra_file(input_path)
            # Sierra map by canonical name
            sierra_map = {row["__canon"]: row for _, row in df.iterrows()}

            # roster (ssn + optional pay/status/type/dept) by canonical name
            roster = self._load_roster()

            # open template
            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            # columns in the template
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

            # letters for formula (if we need to write one)
            reg_L = get_column_letter(reg_col) if reg_col else None
            ot_L  = get_column_letter(ot_col) if ot_col else None
            dt_L  = get_column_letter(dt_col) if dt_col else None

            # fill rows in exact order with canonical matching
            for i, emp in enumerate(order):
                r = start + i
                ws.cell(row=r, column=name_col).value = emp

                k = _canon_name(emp)
                s = sierra_map.get(k)
                ro = roster.get(k, {})

                # SSN / status / type / dept / rate from roster if present
                if ssn_col:  ws.cell(row=r, column=ssn_col).value  = ro.get("ssn", "")
                if stat_col: ws.cell(row=r, column=stat_col).value = ro.get("status", "") or "A"
                if type_col: ws.cell(row=r, column=type_col).value = ro.get("type", "") or "H"
                if dept_col: ws.cell(row=r, column=dept_col).value = ro.get("dept", "")
                if rate_col:
                    try:
                        ws.cell(row=r, column=rate_col).value = float(ro.get("pay_rate", "") or 0.0)
                    except Exception:
                        ws.cell(row=r, column=rate_col).value = 0.0

                # Hours from Sierra
                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0
                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt

                # Row total: if template cell lacks a formula, write =REG+OT+DT
                if tot_col:
                    c = ws.cell(row=r, column=tot_col)
                    has_formula = isinstance(c.value, str) and c.value.startswith("=")
                    if not has_formula and reg_L and ot_L and dt_L:
                        c.value = f"={reg_L}{r}+{ot_L}{r}+{dt_L}{r}"

            wb.save(output_path)

            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return {"success": True, "employees": len(order), "hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
