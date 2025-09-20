# improved_converter.py — FINAL (template-locked, SSN-as-text, Totals$ formula)
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Optional
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

def _load_order(path: Optional[Path] = None) -> List[str]:
    p = path or ORDER_TXT
    if not p.exists():
        return []
    return [ln.strip() for ln in p.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Find the row containing 'Employee Name' (spelled out headers), build a column map.
    This preserves your template’s headings and the A01/A02/… code row below.
    """
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 200) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [_norm(v).replace(" ", "") for v in labels]
        if "employeename" in low:
            for i, v in enumerate(labels, start=1):
                k0 = _norm(v).replace(" ", "")
                if k0 == "employeename": cmap["Employee Name"] = i
                elif k0 in ("ssn","socialsecuritynumber","socialsecurity#","socialsecurityno"): cmap["SSN"] = i
                elif k0 == "status": cmap["Status"] = i
                elif k0 == "type": cmap["Type"] = i
                elif k0 in ("payrate","pay","payrate:"): cmap["Pay Rate"] = i
                elif k0 in ("dept","department"): cmap["Dept"] = i
                elif k0 == "regular": cmap["REGULAR"] = i
                elif k0 in ("overtime","ot"): cmap["OVERTIME"] = i
                elif k0 in ("doubletime","doubletime"): cmap["DOUBLETIME"] = i
                elif k0 == "vacation": cmap["VACATION"] = i
                elif k0 == "sick": cmap["SICK"] = i
                elif k0 == "holiday": cmap["HOLIDAY"] = i
                elif k0 == "bonus": cmap["BONUS"] = i
                elif k0 == "commission": cmap["COMMISSION"] = i
                elif k0 in ("totals","total","sum"): cmap["Totals"] = i
                # daily pink (if present in template)
                elif k0 == "pchrsmon": cmap["AH1"] = i
                elif k0 == "pcttlmon": cmap["AI1"] = i
                elif k0 == "pchrstue": cmap["AH2"] = i
                elif k0 == "pcttltue": cmap["AI2"] = i
                elif k0 == "pchrswed": cmap["AH3"] = i
                elif k0 == "pcttlwed": cmap["AI3"] = i
                elif k0 == "pchrsthu": cmap["AH4"] = i
                elif k0 == "pcttlthu": cmap["AI4"] = i
                elif k0 == "pchrsfri": cmap["AH5"] = i
                elif k0 == "pcttlfri": cmap["AI5"] = i
            return r, cmap
    raise ValueError("Template header row containing 'Employee Name' was not found.")

def _first_data_row(h: int) -> int:
    return h + 1

class SierraToWBSConverter:
    """
    Template-locked converter that:
      • Keeps exact template layout/headers (including the A01/A02/… code row and pink block)
      • Fills SSN/Status/Type/Dept/Pay Rate from roster
      • Fills REG/OT/DT from Sierra
      • Forces SSN as TEXT (keeps leading zeros)
      • Injects Totals$ formula per row (Salary/Commission vs Hourly)
      • Preserves employee order from gold_master_order.txt
    """
    def __init__(self, gold_master_order_path: str | None = None):
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = _load_order(path)

    # ---------- used by /validate ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """Return DF with Name, REGULAR, OVERTIME, DOUBLETIME, Hours, __canon (robust header handling)."""
        def _read(sheet, header):
            try:
                return pd.read_excel(input_path, sheet_name=sheet, header=header).dropna(how="all")
            except Exception:
                return pd.DataFrame()

        df = _read("WEEKLY", 7)
        if df.empty: df = _read(0, 7)
        if df.empty: df = _read(0, 0)
        if df.empty:
            return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME","Hours","__canon"])

        # Remove duplicated header row if present
        if "Employee Name" in df.columns and str(df.iloc[0].get("Employee Name","")).strip() == "Employee Name":
            df = df.iloc[1:].copy()

        name_col = next((c for c in ["Employee Name","Name","Unnamed: 2"] if c in df.columns), None)
        if not name_col:
            return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME","Hours","__canon"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})

        def to_num(series_like):
            return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

        def pick(df_, aliases: List[str]):
            for a in aliases:
                if a in df_.columns:
                    return a
            return None

        reg_c = pick(df, ["REGULAR","Regular","Reg"])
        ot_c  = pick(df, ["OVERTIME","Overtime","OT"])
        dt_c  = pick(df, ["DOUBLETIME","Double Time","DOUBLE TIME","DT"])

        out["REGULAR"]    = to_num(df[reg_c]) if reg_c else 0.0
        out["OVERTIME"]   = to_num(df[ot_c]) if ot_c else 0.0
        out["DOUBLETIME"] = to_num(df[dt_c]) if dt_c else 0.0
        out["Hours"]      = out[["REGULAR","OVERTIME","DOUBLETIME"]].sum(axis=1)

        out = out[out["Name"].astype(str).str.strip().str.len() > 0].copy()
        out["__canon"] = out["Name"].map(_canon_name)
        return out.reset_index(drop=True)

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

        hdr = {c: _norm(c).replace(" ", "") for c in df.columns}

        def col(*aliases):
            al = {a.replace(" ","").lower() for a in aliases}
            for c, k in hdr.items():
                if k in al:
                    return c
            return None

        name_col = col("employeename","name")
        ssn_col  = col("ssn","socialsecuritynumber","socialsecurity#","socialsecurityno")
        stat_col = col("status")
        type_col = col("type")
        dept_col = col("dept","department")
        rate_col = col("payrate","payrate:","pay","pay rate")

        if not name_col:
            print("[WARN] Roster missing Employee Name column.")
            return {}

        out: Dict[str, Dict[str, str]] = {}
        for _, r in df.iterrows():
            nm = str(r.get(name_col, "")).strip()
            if not nm: continue
            k = _canon_name(nm)
            out[k] = {
                "ssn": "" if ssn_col is None or pd.isna(r.get(ssn_col)) else str(r.get(ssn_col)).strip(),
                "status": "" if stat_col is None or pd.isna(r.get(stat_col)) else str(r.get(stat_col)).strip(),
                "type": "" if type_col is None or pd.isna(r.get(type_col)) else str(r.get(type_col)).strip(),
                "dept": "" if dept_col is None or pd.isna(r.get(dept_col)) else str(r.get(dept_col)).strip(),
                "pay_rate": "" if rate_col is None or pd.isna(r.get(rate_col)) else str(r.get(rate_col)).strip(),
            }
        return out

    def convert(self, input_path: str, output_path: str) -> Dict:
        """
        Write into data/wbs_template.xlsx sheet WEEKLY:
          • Names in gold order
          • SSN/Status/Type/Dept/Rate from roster
          • REG/OT/DT from Sierra
          • SSN as text (leading zeros kept)
          • Totals$ formula per row (Salary/Commission vs Hourly)
          • Prefill pink region with zeros (never empty)
        """
        try:
            order = self.gold_master_order or _load_order()
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}

            sierra_df = self.parse_sierra_file(input_path)
            sierra_map = {row["__canon"]: row for _, row in sierra_df.iterrows()}

            roster = self._load_roster()

            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

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
            vac_col  = cmap.get("VACATION")
            sick_col = cmap.get("SICK")
            hol_col  = cmap.get("HOLIDAY")
            bon_col  = cmap.get("BONUS")
            com_col  = cmap.get("COMMISSION")
            tot_col  = cmap.get("Totals")

            # Pink block columns (if present). We’ll prefill zeros so it’s never blank.
            code_row_idx = h + 1
            code_row_vals = [str(c.value).strip().upper() if c.value is not None else "" for c in ws[code_row_idx]]
            pink_cols = []
            for idx, label in enumerate(code_row_vals, start=1):
                if label in {"AH1","AI1","AH2","AI2","AH3","AI3","AH4","AI4","AH5","AI5","ATE"}:
                    pink_cols.append(idx)

            # Prefill zeros in data area of pink block
            max_rows_to_write = max(len(order), 120)
            for r in range(start, start + max_rows_to_write):
                for cidx in pink_cols:
                    c = ws.cell(row=r, column=cidx)
                    if c.value in (None, ""):
                        c.value = 0

            # Column letters for formulas
            rateL = get_column_letter(rate_col) if rate_col else None
            typeL = get_column_letter(type_col) if type_col else None
            regL  = get_column_letter(reg_col)  if reg_col  else None
            otL   = get_column_letter(ot_col)   if ot_col   else None
            dtL   = get_column_letter(dt_col)   if dt_col   else None

            matches = 0
            for i, emp in enumerate(order):
                r = start + i
                ws.cell(row=r, column=name_col).value = emp

                k = _canon_name(emp)
                s = sierra_map.get(k)
                ro = roster.get(k, {})

                # SSN as TEXT to preserve leading zeros
                if ssn_col:
                    ssn_val = (ro.get("ssn", "") or "").strip()
                    c = ws.cell(row=r, column=ssn_col)
                    c.number_format = '@'
                    c.value = ssn_val

                if stat_col: ws.cell(row=r, column=stat_col).value = (ro.get("status","") or "A")
                if type_col: ws.cell(row=r, column=type_col).value = (ro.get("type","") or "H")
                if dept_col: ws.cell(row=r, column=dept_col).value = ro.get("dept","")

                if rate_col:
                    try:
                        val = float(ro.get("pay_rate","") or 0.0)
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

                # Zero-out optional accrual/bonus/commission if cell empty
                for cidx in [vac_col, sick_col, hol_col, bon_col, com_col]:
                    if cidx:
                        cell = ws.cell(row=r, column=cidx)
                        if cell.value in (None, ""):
                            cell.value = 0

                # Inject Totals$ formula per row:
                # IF(Type="S", PayRate, IF(Type="C", PayRate, PayRate*(REG + 1.5*OT + 2*DT)))
                if tot_col and rateL and typeL and regL and otL and dtL:
                    formula = (
                        f'=IF(UPPER({typeL}{r})="S",{rateL}{r},'
                        f'IF(UPPER({typeL}{r})="C",{rateL}{r},'
                        f'{rateL}{r}*({regL}{r}+1.5*{otL}{r}+2*{dtL}{r})))'
                    )
                    ws.cell(row=r, column=tot_col).value = formula

                if s is not None:
                    matches += 1

            wb.save(output_path)

            total_hours = float(sierra_df[["REGULAR","OVERTIME","DOUBLETIME"]].sum().sum()) if not sierra_df.empty else 0.0
            if matches == 0:
                print("[WARN] No Sierra rows matched names in gold order after canonicalization.")

            return {"success": True, "employees": len(order), "hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
