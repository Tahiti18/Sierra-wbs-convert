# improved_converter.py — FINAL
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

# ---------------- helpers ----------------
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
        if v is None or (isinstance(v, str) and v.strip() == ""):
            return 0.0
        return float(v)
    except Exception:
        return 0.0

def _load_order(default_path: Path) -> List[str]:
    if not default_path.exists():
        return []
    return [ln.strip() for ln in default_path.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Locate the real human headers row (with 'Employee Name').
    Return (row_index, {Header->column_index})
    """
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 150) + 1):
        vals = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        joined = "".join(vals).lower()
        if "employee name" in joined:
            # Normalize the key map
            for i, v in enumerate(vals, start=1):
                k = v.strip().lower().replace(" ", "")
                if k == "employeename":
                    cmap["Employee Name"] = i
                elif k in ("ssn","socialsecuritynumber","socialsecurity#","socialsecurityno"):
                    cmap["SSN"] = i
                elif k == "regular":
                    cmap["REGULAR"] = i
                elif k in ("overtime","ot"):
                    cmap["OVERTIME"] = i
                elif k in ("doubletime","double time"):
                    cmap["DOUBLETIME"] = i
                elif k == "status":
                    cmap["Status"] = i
                elif k == "type":
                    cmap["Type"] = i
                elif k in ("payrate","pay","payrate"):
                    cmap["Pay Rate"] = i
                elif k in ("dept","department"):
                    cmap["Dept"] = i
                elif k in ("totals","total","sum"):
                    cmap["Totals"] = i
                elif k == "pchrsmon":
                    cmap["AH1"] = i
                elif k == "pcttlmon":
                    cmap["AI1"] = i
                elif k == "pchrstue":
                    cmap["AH2"] = i
                elif k == "pcttltue":
                    cmap["AI2"] = i
                elif k == "pchrswed":
                    cmap["AH3"] = i
                elif k == "pcttlwed":
                    cmap["AI3"] = i
                elif k == "pchrsthu":
                    cmap["AH4"] = i
                elif k == "pcttlthu":
                    cmap["AI4"] = i
                elif k == "pchrsfri":
                    cmap["AH5"] = i
                elif k == "pcttlfri":
                    cmap["AI5"] = i
            return r, cmap
    raise ValueError("Could not locate header row containing 'Employee Name'")

def _first_data_row(h: int) -> int:
    return h + 1

# ---------------- converter ----------------
class SierraToWBSConverter:
    """
    Produces data rows inside data/wbs_template.xlsx, preserving all formatting and formulas.
    - Name/SSN/Status/Type/Dept/Pay Rate from gold roster
    - REG/OT/DT from Sierra
    - Daily pink columns populated if Sierra provides day columns; else left as 0
    """
    def __init__(self, gold_master_order_path: str | None = None):
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = _load_order(path)

    # ---------- used by /validate ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Return DF with:
            Name, REGULAR, OVERTIME, DOUBLETIME, Hours, __canon
        Tries WEEKLY @ header row 8 first (to match the gold layout), then falls back.
        Also tries per-day columns if present (Mon..Fri).
        """
        def _read(sheet, header):
            try:
                return pd.read_excel(input_path, sheet_name=sheet, header=header).dropna(how="all")
            except Exception:
                return pd.DataFrame()

        df = _read("WEEKLY", 7)
        if df.empty:
            df = _read(0, 7)
        if df.empty:
            df = _read(0, 0)

        # Name column
        name_col = None
        for cand in ["Employee Name", "Name", "Unnamed: 2"]:
            if cand in df.columns:
                name_col = cand
                break
        if not name_col:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})

        # REG / OT / DT
        def to_num(series_like):
            return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

        colmap = {
            "REGULAR": ["REGULAR", "Regular"],
            "OVERTIME": ["OVERTIME", "Overtime", "OT"],
            "DOUBLETIME": ["DOUBLETIME", "Double Time", "DOUBLE TIME"],
        }
        for dst, srcs in colmap.items():
            src = next((c for c in srcs if c in df.columns), None)
            out[dst] = to_num(df[src]) if src else 0.0

        # Optional per-day columns (pink side)
        # We only collect if present to fill template; otherwise left as zeros.
        day_aliases = {
            "MON": ["Mon", "MON", "Monday"],
            "TUE": ["Tue", "TUE", "Tuesday"],
            "WED": ["Wed", "WED", "Wednesday"],
            "THU": ["Thu", "THU", "Thursday"],
            "FRI": ["Fri", "FRI", "Friday"],
        }
        for key, aliases in day_aliases.items():
            src = next((c for c in aliases if c in df.columns), None)
            out[f"PC_HRS_{key}"] = to_num(df[src]) if src else 0.0
            # Totals (money) typically a separate column; if not present keep 0
            amt_src = next((c for c in [f"{key} Total", f"{key} Amount", f"{key}_TOTAL"] if c in df.columns), None)
            out[f"PC_TTL_{key}"] = to_num(df[amt_src]) if amt_src else 0.0

        out = out[out["Name"].astype(str).str.strip().str.len() > 0].copy()
        out["Hours"] = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out["__canon"] = out["Name"].map(_canon_name)
        return out

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        """Return canonical_name -> {ssn,status,type,dept,pay_rate} from /data/gold_master_roster.csv"""
        if not ROSTER_CSV.exists():
            print(f"[WARN] Roster not found: {ROSTER_CSV}")
            return {}
        try:
            df = pd.read_csv(ROSTER_CSV)
        except Exception as e:
            print(f"[WARN] Failed to read roster CSV: {e}")
            return {}

        # header aliasing
        header_norm = {c: _norm(c).replace(" ", "") for c in df.columns}

        def pick(*aliases):
            al = {a.replace(" ", "").lower() for a in aliases}
            for c, n in header_norm.items():
                if n in al:
                    return c
            return None

        name_col = pick("employeename", "name")
        ssn_col  = pick("ssn","socialsecuritynumber","socialsecurity#")
        stat_col = pick("status")
        type_col = pick("type")
        dept_col = pick("dept","department")
        rate_col = pick("payrate","payrate","pay","pay rate")

        if not name_col:
            print("[WARN] Roster missing an Employee Name column.")
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
        return out

    # ---------- main convert ----------
    def convert(self, input_path: str, output_path: str) -> Dict:
        try:
            # order
            order = self.gold_master_order[:] or _load_order(ORDER_TXT)
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}

            # sierra
            df = self.parse_sierra_file(input_path)
            sierra_map = {row["__canon"]: row for _, row in df.iterrows()}

            # roster
            roster = self._load_roster()

            # template
            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            # column map
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

            # Pink side (daily piecework columns) if present in template
            AH1 = cmap.get("AH1"); AI1 = cmap.get("AI1")
            AH2 = cmap.get("AH2"); AI2 = cmap.get("AI2")
            AH3 = cmap.get("AH3"); AI3 = cmap.get("AI3")
            AH4 = cmap.get("AH4"); AI4 = cmap.get("AI4")
            AH5 = cmap.get("AH5"); AI5 = cmap.get("AI5")

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

                # roster fields
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

                # Daily (pink) if template columns exist
                if s is not None:
                    # Hours
                    if AH1: ws.cell(row=r, column=AH1).value = _num(s.get("PC_HRS_MON", 0))
                    if AH2: ws.cell(row=r, column=AH2).value = _num(s.get("PC_HRS_TUE", 0))
                    if AH3: ws.cell(row=r, column=AH3).value = _num(s.get("PC_HRS_WED", 0))
                    if AH4: ws.cell(row=r, column=AH4).value = _num(s.get("PC_HRS_THU", 0))
                    if AH5: ws.cell(row=r, column=AH5).value = _num(s.get("PC_HRS_FRI", 0))
                    # Amounts
                    if AI1: ws.cell(row=r, column=AI1).value = _num(s.get("PC_TTL_MON", 0))
                    if AI2: ws.cell(row=r, column=AI2).value = _num(s.get("PC_TTL_TUE", 0))
                    if AI3: ws.cell(row=r, column=AI3).value = _num(s.get("PC_TTL_WED", 0))
                    if AI4: ws.cell(row=r, column=AI4).value = _num(s.get("PC_TTL_THU", 0))
                    if AI5: ws.cell(row=r, column=AI5).value = _num(s.get("PC_TTL_FRI", 0))

                # Row total formula safeguard (=REG+OT+DT) if template cell has no formula
                if tot_col and regL and otL and dtL:
                    c = ws.cell(row=r, column=tot_col)
                    if not (isinstance(c.value, str) and c.value.startswith("=")):
                        c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

            wb.save(output_path)

            total_hours = float(df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not df.empty else 0.0
            return {"success": True, "employees": len(order), "hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
