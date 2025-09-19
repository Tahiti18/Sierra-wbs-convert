# improved_converter.py  — stable with current main.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

# repo paths
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"  # in the template

def _norm(s: str) -> str:
    return (s or "").strip().lower().replace(" ", "")

def _num(v) -> float:
    try:
        if v is None or (isinstance(v, str) and v.strip() == ""):
            return 0.0
        return float(v)
    except Exception:
        return 0.0

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Find the row containing 'Employee Name' and build a col map by header text.
    """
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 120) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [_norm(v) for v in labels]
        if "employeename" in low:
            for i, v in enumerate(labels, start=1):
                k = _norm(v)
                if k in ("employeename", "employee name"):
                    cmap["Employee Name"] = i
                elif k in ("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno"):
                    cmap["SSN"] = i
                elif k == "regular":
                    cmap["REGULAR"] = i
                elif k in ("overtime", "ot"):
                    cmap["OVERTIME"] = i
                elif k in ("doubletime", "double time"):
                    cmap["DOUBLETIME"] = i
                elif k in ("totals", "total", "sum"):
                    cmap["Totals"] = i
            return r, cmap
    raise ValueError("Could not locate header row containing 'Employee Name'")

def _first_data_row(h: int) -> int:
    return h + 1

class SierraToWBSConverter:
    """
    Opens data/wbs_template.xlsx and fills Name, SSN, REG/OT/DT
    using gold order + roster. Keeps headers/format/totals.
    """

    def __init__(self, gold_master_order_path: str | None = None):
        # keep attribute for main.py logging (prevents AttributeError)
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = [
                ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()
            ]

    # ---------- used by /validate ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Return DataFrame with columns: Name, REGULAR, OVERTIME, DOUBLETIME, Hours.
        Auto-detects the Sierra sheet: tries 'WEEKLY', else first sheet.
        Tries header at row 8; if columns missing, falls back to header row 1.
        """
        # 1) pick sheet safely
        try:
            df = pd.read_excel(input_path, sheet_name="WEEKLY", header=7)  # common case
        except Exception:
            df = pd.read_excel(input_path, sheet_name=0, header=7)  # first sheet

        df = df.dropna(how="all")

        # if the first data row is actually a duplicated header, drop it
        if not df.empty and "Employee Name" in df.columns and str(df.iloc[0]["Employee Name"]).strip() == "Employee Name":
            df = df.iloc[1:]

        # If expected columns missing, retry with header=0
        expected = {"Employee Name", "REGULAR", "OVERTIME", "DOUBLETIME"}
        if len(expected.intersection(set(map(str, df.columns)))) < 2:
            try:
                df = pd.read_excel(input_path, sheet_name=0, header=0)
                df = df.dropna(how="all")
            except Exception:
                pass

        # choose name column
        name_col = None
        for cand in ["Employee Name", "Name", "Unnamed: 2"]:
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

    def _load_roster_ssn(self) -> Dict[str, str]:
        if not ROSTER_CSV.exists():
            return {}
        df = pd.read_csv(ROSTER_CSV)
        cols = {c: _norm(c) for c in df.columns}
        name_col = next((c for c, n in cols.items() if n in ("employeename", "name")), None)
        ssn_col = next((c for c, n in cols.items() if n == "ssn"), None)
        if not name_col or not ssn_col:
            return {}
        out: Dict[str, str] = {}
        for _, r in df[[name_col, ssn_col]].iterrows():
            nm = str(r[name_col]).strip()
            sv = "" if pd.isna(r[ssn_col]) else str(r[ssn_col]).strip()
            if nm:
                out[nm] = sv
        return out

    # ---------- main convert ----------
    def convert(self, input_path: str, output_path: str) -> Dict:
        try:
            # gold order
            order: List[str] = self.gold_master_order[:]
            if not order and ORDER_TXT.exists():
                order = [
                    ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines() if ln.strip()
                ]
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}

            # values from Sierra
            df = self.parse_sierra_file(input_path)
            sierra_map = {r["Name"]: r for _, r in df.iterrows()}

            # SSNs
            ssn_map = self._load_roster_ssn()

            # open template
            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            # map columns from template headers (do NOT change headers)
            h, cmap = _find_header_row(ws)
            start = _first_data_row(h)
            name_col = cmap["Employee Name"]
            ssn_col  = cmap.get("SSN")
            reg_col  = cmap.get("REGULAR")
            ot_col   = cmap.get("OVERTIME")
            dt_col   = cmap.get("DOUBLETIME")
            tot_col  = cmap.get("Totals")

            # fill rows in exact gold order
            for i, emp in enumerate(order):
                r = start + i
                ws.cell(row=r, column=name_col).value = emp
                if ssn_col:
                    ws.cell(row=r, column=ssn_col).value = ssn_map.get(emp, "")

                s = sierra_map.get(emp)
                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0

                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt

                # totals only if template cell is not a formula
                if tot_col:
                    c = ws.cell(row=r, column=tot_col)
                    if not (isinstance(c.value, str) and str(c.value).startswith("=")):
                        c.value = reg + ot + dt

            wb.save(output_path)

            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return {"success": True, "employees": len(order), "hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
