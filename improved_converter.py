# improved_converter.py (REPLACE THE WHOLE FILE WITH THIS)
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

def _norm(h: str) -> str:
    return (h or "").strip().lower().replace(" ", "")

def _load_order() -> List[str]:
    if not ORDER_TXT.exists():
        return []
    return [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _load_roster_ssn() -> Dict[str, str]:
    if not ROSTER_CSV.exists():
        return {}
    df = pd.read_csv(ROSTER_CSV)
    cols = {c: _norm(c) for c in df.columns}
    name_col = next((c for c,n in cols.items() if n in ("employeename","name")), None)
    ssn_col  = next((c for c,n in cols.items() if n == "ssn"), None)
    if not name_col or not ssn_col:
        return {}
    out = {}
    for _, r in df[[name_col, ssn_col]].iterrows():
        nm = str(r[name_col]).strip()
        sv = "" if pd.isna(r[ssn_col]) else str(r[ssn_col]).strip()
        if nm:
            out[nm] = sv
    return out

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str,int]]:
    cmap: Dict[str,int] = {}
    for r in range(1, min(ws.max_row, 80) + 1):
        vals = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [_norm(v) for v in vals]
        if "employeename" in low:
            for i,v in enumerate(vals, start=1):
                k = _norm(v)
                if k in ("employeename","employee name"): cmap["Employee Name"]=i
                elif k in ("ssn","socialsecuritynumber","socialsecurity#"): cmap["SSN"]=i
                elif k in ("regular",): cmap["REGULAR"]=i
                elif k in ("overtime","ot"): cmap["OVERTIME"]=i
                elif k in ("doubletime","doubletime", "double time"): cmap["DOUBLETIME"]=i
                elif k in ("totals","total","sum"): cmap["Totals"]=i
            return r, cmap
    raise ValueError("Header row with 'Employee Name' not found in template")

def _first_data_row(header_row: int) -> int:
    return header_row + 1

def _num(v) -> float:
    try:
        if v is None or (isinstance(v,str) and v.strip()==""): return 0.0
        return float(v)
    except Exception:
        return 0.0

class SierraToWBSConverter:
    def __init__(self, _unused=None):
        pass

    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        # Sierra weekly sheet: header around row 8; handle variants safely
        df = pd.read_excel(input_path, sheet_name="WEEKLY", header=7)
        df = df.dropna(how="all")
        # drop duplicated header row if present
        if (df.iloc[0:1].astype(str).apply(lambda x: (x == 'Employee Name').any(), axis=1).any()):
            df = df.iloc[1:]
        # locate name
        name_col = None
        for cand in ["Employee Name","Unnamed: 2","Name"]:
            if cand in df.columns:
                name_col = cand; break
        if not name_col:
            return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME"])
        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})
        for col in ["REGULAR","OVERTIME","DOUBLETIME"]:
            out[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0.0)
        out = out[out["Name"].str.len() > 0]
        return out

    def convert(self, input_path: str, output_path: str) -> Dict:
        try:
            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            order = _load_order()
            if not order:
                return {"success": False, "error": f"gold_master_order.txt missing/empty"}
            ssn_map = _load_roster_ssn()
            df = self.parse_sierra_file(input_path)

            # open template
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            # map columns from the template
            h, cmap = _find_header_row(ws)
            data_start = _first_data_row(h)
            name_col = cmap["Employee Name"]
            ssn_col  = cmap.get("SSN")
            reg_col  = cmap.get("REGULAR")
            ot_col   = cmap.get("OVERTIME")
            dt_col   = cmap.get("DOUBLETIME")
            tot_col  = cmap.get("Totals")

            # name -> values from Sierra
            s_map = {r["Name"]: r for _, r in df.iterrows()}

            # fill rows in exact gold order
            for i, emp in enumerate(order):
                r = data_start + i
                # DO NOT touch header/totals rows; only write row values
                ws.cell(row=r, column=name_col).value = emp
                if ssn_col: ws.cell(row=r, column=ssn_col).value = ssn_map.get(emp, "")

                srow = s_map.get(emp)
                reg = _num(srow["REGULAR"]) if srow is not None else 0.0
                ot  = _num(srow["OVERTIME"]) if srow is not None else 0.0
                dt  = _num(srow["DOUBLETIME"]) if srow is not None else 0.0

                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt

                # totals: only write if the template cell is NOT a formula
                if tot_col:
                    c = ws.cell(row=r, column=tot_col)
                    if not (isinstance(c.value, str) and str(c.value).startswith("=")):
                        c.value = reg + ot + dt

            wb.save(output_path)

            total_hours = float(df[["REGULAR","OVERTIME","DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return {"success": True, "employees": len(order), "hours": total_hours}
        except Exception as e:
            return {"success": False, "error": str(e)}
