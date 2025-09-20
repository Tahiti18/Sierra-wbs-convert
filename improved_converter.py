# improved_converter.py – template-driven writer with roster merge and totals column filled
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# ---- paths ----
HERE = Path(__file__).resolve().parent
ROOT = HERE
DATA = (ROOT / "data")
if not DATA.exists():
    DATA = (HERE.parent / "data")
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

def _canon_name(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace(".", "")
    s = re.sub(r"\s*,\s*", ", ", s)
    s = s.replace(" ,", ",")
    return s.lower()

def _to_num(val) -> float:
    try:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return 0.0
        return float(val)
    except Exception:
        return 0.0

def _load_order(p: Path|None=None) -> List[str]:
    src = p or ORDER_TXT
    if not src.exists(): return []
    return [ln.strip() for ln in src.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _read_sierra_table(xlsx_path: str) -> pd.DataFrame:
    # Try WEEKLY header row 8, fallbacks
    tried = [( "WEEKLY",7), (0,7), (0,0)]
    for sheet, hdr in tried:
        try:
            df = pd.read_excel(xlsx_path, sheet_name=sheet, header=hdr).dropna(how='all')
        except Exception:
            continue

        name_col = next((c for c in df.columns if str(c).strip().lower().replace(" ","") in ("employeename","name")), None)
        if not name_col:
            continue

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})
        def pick(*aliases):
            for c in df.columns:
                n = str(c).strip().lower()
                if any(a == n or a in n for a in aliases):
                    return c
            return None

        reg = pick("regular")
        ot  = pick("overtime","ot")
        dt  = pick("double")
        out["REGULAR"]   = pd.to_numeric(df[reg], errors='coerce').fillna(0.0) if reg is not None else 0.0
        out["OVERTIME"]  = pd.to_numeric(df[ot],  errors='coerce').fillna(0.0) if ot  is not None else 0.0
        out["DOUBLETIME"]= pd.to_numeric(df[dt],  errors='coerce').fillna(0.0) if dt  is not None else 0.0

        out = out[out["Name"].str.len() > 0]
        out = out[out["Name"].str.lower() != "employee name"]
        out["__canon"] = out["Name"].map(_canon_name)
        return out.reset_index(drop=True)
    return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME","__canon"])

def _read_roster() -> Dict[str, Dict[str,str]]:
    if not ROSTER_CSV.exists(): return {}
    try:
        df = pd.read_csv(ROSTER_CSV)
    except Exception:
        return {}

    def norm(s): return str(s).strip().lower().replace(" ","")
    aliases = {}
    for c in df.columns:
        aliases[norm(c)] = c

    def pick(*keys):
        for k in keys:
            if k in aliases: return aliases[k]
        return None

    name = pick("employeename","name")
    ssn  = pick("ssn","socialsecuritynumber","socialsecurity#")
    stat = pick("status")
    typ  = pick("type")
    dept = pick("dept","department")
    rate = pick("payrate","payrate","payrate","payrate","payrate") or pick("payrate","payrate") or pick("payrate","pay rate") or pick("rate")

    out: Dict[str, Dict[str,str]] = {}
    if not name: return out
    for _, r in df.iterrows():
        nm = str(r.get(name,"")).strip()
        if not nm: continue
        k = _canon_name(nm)
        out[k] = {
            "ssn": "" if ssn  is None or pd.isna(r.get(ssn))  else str(r.get(ssn)).strip(),
            "status": "" if stat is None or pd.isna(r.get(stat)) else str(r.get(stat)).strip(),
            "type": "" if typ is None or pd.isna(r.get(typ)) else str(r.get(typ)).strip(),
            "dept": "" if dept is None or pd.isna(r.get(dept)) else str(r.get(dept)).strip(),
            "pay_rate": "" if rate is None or pd.isna(r.get(rate)) else str(r.get(rate)).strip(),
        }
    return out

def _find_header(ws: Worksheet) -> Tuple[int, Dict[str,int]]:
    """
    Find the row that contains the printable headers (the row with 'Employee Name')
    and return a column map for SSN/Employee Name/Status/Type/Pay Rate/Dept/REG/OT/DT/Totals.
    """
    for r in range(1, min(ws.max_row, 150)+1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [x.lower().replace(" ","") for x in labels]
        if "employeename" in low:
            cmap: Dict[str,int] = {}
            for i, v in enumerate(labels, start=1):
                vv = v.lower().strip()
                key = vv.replace(" ", "")
                if key in ("ssn","socialsecuritynumber","socialsecurity#"): cmap["SSN"]=i
                elif key=="employeename": cmap["Employee Name"]=i
                elif key=="status": cmap["Status"]=i
                elif key=="type": cmap["Type"]=i
                elif key in ("payrate","pay rate","pay"): cmap["Pay Rate"]=i
                elif key in ("dept","department"): cmap["Dept"]=i
                elif key=="regular": cmap["REG"]=i
                elif key in ("overtime","ot"): cmap["OT"]=i
                elif "double" in key: cmap["DT"]=i
                elif key in ("totals","total","sum"): cmap["Totals"]=i
            return r, cmap
    raise RuntimeError("Could not locate header row containing 'Employee Name'")

class SierraToWBSConverter:
    def __init__(self, gold_master_order_path: str|None=None):
        p = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        self.gold_master_order: List[str] = _load_order(p)

    def convert(self, input_path: str, output_path: str) -> Dict:
        # sanity checks
        order = self.gold_master_order[:] or _load_order()
        if not order:
            return {"success": False, "error": "gold_master_order.txt missing/empty"}

        if not TEMPLATE_XLSX.exists():
            return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}

        sierra = _read_sierra_table(input_path)
        roster = _read_roster()

        wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
        if TARGET_SHEET not in wb.sheetnames:
            return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
        ws: Worksheet = wb[TARGET_SHEET]

        hdr_row, cmap = _find_header(ws)
        start = hdr_row + 1

        name_c = cmap.get("Employee Name")
        ssn_c  = cmap.get("SSN")
        stat_c = cmap.get("Status")
        type_c = cmap.get("Type")
        rate_c = cmap.get("Pay Rate")
        dept_c = cmap.get("Dept")
        reg_c  = cmap.get("REG")
        ot_c   = cmap.get("OT")
        dt_c   = cmap.get("DT")
        tot_c  = cmap.get("Totals")  # rightmost pink “Totals” column

        regL = get_column_letter(reg_c) if reg_c else None
        otL  = get_column_letter(ot_c) if ot_c else None
        dtL  = get_column_letter(dt_c) if dt_c else None

        # quick map for hours by canonical name
        s_map = {row["__canon"]: row for _, row in sierra.iterrows()}

        filled = 0
        for i, emp in enumerate(order):
            r = start + i
            ws.cell(row=r, column=name_c).value = emp

            key = _canon_name(emp)
            s   = s_map.get(key)
            ro  = roster.get(key, {})

            if ssn_c:  ws.cell(row=r, column=ssn_c).value  = ro.get("ssn", "")
            if stat_c: ws.cell(row=r, column=stat_c).value = ro.get("status", "") or "A"
            if type_c: ws.cell(row=r, column=type_c).value = ro.get("type", "") or "H"
            if dept_c: ws.cell(row=r, column=dept_c).value = ro.get("dept", "")

            if rate_c:
                try:
                    ws.cell(row=r, column=rate_c).value = float(ro.get("pay_rate","") or 0.0)
                except Exception:
                    ws.cell(row=r, column=rate_c).value = 0.0

            reg = _to_num(s["REGULAR"])    if s is not None else 0.0
            ot  = _to_num(s["OVERTIME"])   if s is not None else 0.0
            dt  = _to_num(s["DOUBLETIME"]) if s is not None else 0.0
            if reg_c: ws.cell(row=r, column=reg_c).value = reg
            if ot_c:  ws.cell(row=r, column=ot_c).value  = ot
            if dt_c:  ws.cell(row=r, column=dt_c).value  = dt

            # Ensure the pink far-right Totals cell has a formula =REG+OT+DT if it doesn't already
            if tot_c and regL and otL and dtL:
                c = ws.cell(row=r, column=tot_c)
                if not (isinstance(c.value, str) and c.value.startswith("=")):
                    c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

            if s is not None:
                filled += 1

        wb.save(output_path)

        total_hours = float(sierra[["REGULAR","OVERTIME","DOUBLETIME"]].sum(numeric_only=True).sum()) if not sierra.empty else 0.0
        return {"success": True, "employees": len(order), "hours": total_hours}
