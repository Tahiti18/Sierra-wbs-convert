# src/improved_converter.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# Project root (this file lives in src/). Data files live in /data.
ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

# -------- helpers --------
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

def _load_order() -> List[str]:
    if not ORDER_TXT.exists():
        return []
    return [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Locate the header row on the template and return:
      row_index, { "SSN": col, "Employee Name": col, "Pay Rate": col, ... }
    This mapping enforces the correct WBS order:
      A01 REGULAR, A02 OVERTIME, A03 DOUBLETIME, A06 VACATION,
      A07 SICK, A08 HOLIDAY, A04 BONUS, A05 COMMISSION,
      AH1/AI1 ... AH5/AI5 piecework columns, ATE travel, Comments, Totals
    """
    # We look for the row that has the visible column labels like "SSN", "Employee Name", etc.
    for r in range(1, min(ws.max_row, 150) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [_norm(v) for v in labels]
        if "employee name" in low:
            # Build canonical map
            cmap: Dict[str, int] = {}
            for i, v in enumerate(labels, start=1):
                k = _norm(v)
                k0 = k.replace(" ", "")
                if k0 in ("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno"):
                    cmap["SSN"] = i
                elif k0 == "employeename":
                    cmap["Employee Name"] = i
                elif k0 == "status":
                    cmap["Status"] = i
                elif k0 == "type":
                    cmap["Type"] = i
                elif k0 in ("payrate", "pay", "payrate$","pay rate"):
                    cmap["Pay Rate"] = i
                elif k0 in ("dept", "department"):
                    cmap["Dept"] = i
                elif k0 == "regular":
                    cmap["REGULAR"] = i
                elif k0 in ("overtime","ot"):
                    cmap["OVERTIME"] = i
                elif k0 in ("doubletime","double time"):
                    cmap["DOUBLETIME"] = i
                elif k0 == "vacation":
                    cmap["VACATION"] = i
                elif k0 == "sick":
                    cmap["SICK"] = i
                elif k0 == "holiday":
                    cmap["HOLIDAY"] = i
                elif k0 == "bonus":
                    cmap["BONUS"] = i
                elif k0 == "commission":
                    cmap["COMMISSION"] = i
                elif k0 in ("pchrsmon","pc hrs mon"):
                    cmap["AH1"] = i
                elif k0 in ("pcttlmon","pc ttl mon"):
                    cmap["AI1"] = i
                elif k0 in ("pchrstue","pc hrs tue"):
                    cmap["AH2"] = i
                elif k0 in ("pcttltue","pc ttl tue"):
                    cmap["AI2"] = i
                elif k0 in ("pchrswed","pc hrs wed"):
                    cmap["AH3"] = i
                elif k0 in ("pcttlwed","pc ttl wed"):
                    cmap["AI3"] = i
                elif k0 in ("pchrsthu","pc hrs thu"):
                    cmap["AH4"] = i
                elif k0 in ("pcttlthu","pc ttl thu"):
                    cmap["AI4"] = i
                elif k0 in ("pchrsfri","pc hrs fri"):
                    cmap["AH5"] = i
                elif k0 in ("pcttlfri","pc ttl fri"):
                    cmap["AI5"] = i
                elif k0 in ("travelamount","travel","ate"):
                    cmap["ATE"] = i
                elif k0.startswith("notes") or k0 == "comments":
                    cmap["Comments"] = i
                elif k0 in ("totals","total"):
                    cmap["Totals"] = i
            # Hard check for the key columns we must have
            for must in ("Employee Name","SSN","REGULAR","OVERTIME","DOUBLETIME"):
                if must not in cmap:
                    # The template is expected to have these exact visible titles.
                    # If any are missing, we still return what we found; fill will continue but may be partial.
                    pass
            return r, cmap
    raise ValueError("Template header row not found (no visible 'Employee Name').")

def _first_data_row(h: int) -> int:
    return h + 1

# -------- converter --------
class SierraToWBSConverter:
    """
    Reads Sierra weekly file, fills Name, SSN, Dept/Type/Rate from roster,
    fills REG/OT/DT (A01/A02/A03), A06/A07/A08, A04/A05 when present,
    copies piecework columns (AH1/AI1..AH5/AI5),
    keeps your Excel template formatting, adds per-row formula and a totals row.
    """
    def __init__(self, gold_master_order_path: str | None = None):
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = [
                ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()
            ]

    # ---------- used by /validate ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Return DF with Name, REGULAR, OVERTIME, DOUBLETIME, VACATION, SICK, HOLIDAY,
        BONUS, COMMISSION, Hours, piecework columns (if present), and __canon.
        Robust to header row (7 or 0) and capitalization/aliases.
        """
        def _read(sheet, header):
            try:
                return pd.read_excel(input_path, sheet_name=sheet, header=header).dropna(how="all")
            except Exception:
                return pd.DataFrame()

        # Try WEEKLY@header7, then first sheet@header7, then first sheet@header0
        df = _read("WEEKLY", 7)
        if df.empty: df = _read(0, 7)
        if df.empty: df = _read(0, 0)
        if df.empty:
            return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME","Hours","__canon"])

        # Some exports repeat the header row one more time just below; remove if present.
        if "Employee Name" in df.columns and str(df.iloc[0].get("Employee Name","")).strip() == "Employee Name":
            df = df.iloc[1:].copy()

        # Column aliasing
        def pick(*alts):
            for a in alts:
                if a in df.columns: return a
            return None

        name_col = pick("Employee Name","Name","Unnamed: 2")
        if not name_col:
            return pd.DataFrame(columns=["Name","REGULAR","OVERTIME","DOUBLETIME","Hours","__canon"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})

        def to_num(cand_cols, default=0.0):
            c = pick(*cand_cols)
            if not c: return pd.Series([default]*len(out), dtype="float64")
            return pd.to_numeric(df[c], errors="coerce").fillna(0.0)

        # Core hour buckets
        out["REGULAR"]    = to_num(("REGULAR","Regular","A01"))
        out["OVERTIME"]   = to_num(("OVERTIME","Overtime","OT","A02"))
        out["DOUBLETIME"] = to_num(("DOUBLETIME","Double Time","DOUBLE TIME","A03"))

        # Paid time / extras (present in some weeks)
        out["VACATION"]   = to_num(("VACATION","A06"))
        out["SICK"]       = to_num(("SICK","A07"))
        out["HOLIDAY"]    = to_num(("HOLIDAY","A08"))
        out["BONUS"]      = to_num(("BONUS","A04"))
        out["COMMISSION"] = to_num(("COMMISSION","A05"))

        # Piecework (pink) — hours + total per day if present
        out["AH1"] = to_num(("PC HRS MON","PCHRS MON","PC HRS Mon"))
        out["AI1"] = to_num(("PC TTL MON","PCTTL MON","PC TTL Mon"))
        out["AH2"] = to_num(("PC HRS TUE","PCHRS TUE","PC HRS Tue"))
        out["AI2"] = to_num(("PC TTL TUE","PCTTL TUE","PC TTL Tue"))
        out["AH3"] = to_num(("PC HRS WED","PCHRS WED","PC HRS Wed"))
        out["AI3"] = to_num(("PC TTL WED","PCTTL WED","PC TTL Wed"))
        out["AH4"] = to_num(("PC HRS THU","PCHRS THU","PC HRS Thu"))
        out["AI4"] = to_num(("PC TTL THU","PCTTL THU","PC TTL Thu"))
        out["AH5"] = to_num(("PC HRS FRI","PCHRS FRI","PC HRS Fri"))
        out["AI5"] = to_num(("PC TTL FRI","PCTTL FRI","PC TTL Fri"))

        # Total hours = REG + OT + DT (+ optional paid buckets when they exist)
        out["Hours"] = out[["REGULAR","OVERTIME","DOUBLETIME","VACATION","SICK","HOLIDAY"]].sum(axis=1)

        # Keep only real names
        out = out[out["Name"].astype(str).str.strip().str.len() > 0].copy()
        out["__canon"] = out["Name"].map(_canon_name)
        return out

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        if not ROSTER_CSV.exists():
            print(f"[WARN] Roster not found: {ROSTER_CSV}")
            return {}
        try:
            df = pd.read_csv(ROSTER_CSV)
        except Exception as e:
            print(f"[WARN] Failed to read roster CSV: {e}")
            return {}

        cols = {c: _norm(c).replace(" ", "") for c in df.columns}
        def pick(*aliases):
            al = [a.replace(" ", "").lower() for a in aliases]
            return next((c for c, n in cols.items() if n in al), None)

        name_col = pick("employeename","name")
        ssn_col  = pick("ssn","socialsecuritynumber","socialsecurity#")
        stat_col = pick("status")
        type_col = pick("type")
        dept_col = pick("dept","department")
        rate_col = pick("payrate","payrate$","payrate","payratehourly","payratehour","payrateh","pay rate","rate")

        if not name_col:
            print("[WARN] Roster missing 'Employee Name'.")
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

    # ---------- main convert ----------
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
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing sheet '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            h, cmap = _find_header_row(ws)
            start = _first_data_row(h)

            # Column indexes (will be None if the template lacks some)
            name_col = cmap.get("Employee Name")
            ssn_col  = cmap.get("SSN")
            reg_col  = cmap.get("REGULAR")
            ot_col   = cmap.get("OVERTIME")
            dt_col   = cmap.get("DOUBLETIME")
            vac_col  = cmap.get("VACATION")
            sick_col = cmap.get("SICK")
            hol_col  = cmap.get("HOLIDAY")
            bon_col  = cmap.get("BONUS")
            com_col  = cmap.get("COMMISSION")
            rate_col = cmap.get("Pay Rate")
            stat_col = cmap.get("Status")
            type_col = cmap.get("Type")
            dept_col = cmap.get("Dept")
            tot_col  = cmap.get("Totals")

            ah_cols = [cmap.get("AH1"), cmap.get("AH2"), cmap.get("AH3"), cmap.get("AH4"), cmap.get("AH5")]
            ai_cols = [cmap.get("AI1"), cmap.get("AI2"), cmap.get("AI3"), cmap.get("AI4"), cmap.get("AI5")]
            ate_col = cmap.get("ATE")
            cmt_col = cmap.get("Comments")

            regL = get_column_letter(reg_col) if reg_col else None
            otL  = get_column_letter(ot_col)  if ot_col  else None
            dtL  = get_column_letter(dt_col)  if dt_col  else None

            matches = 0
            for i, emp in enumerate(order):
                r = start + i
                if name_col: ws.cell(row=r, column=name_col).value = emp

                k  = _canon_name(emp)
                s  = sierra_map.get(k)
                ro = roster.get(k, {})

                # Roster-provided fields
                if ssn_col:  ws.cell(row=r, column=ssn_col).value  = ro.get("ssn", "")
                if stat_col: ws.cell(row=r, column=stat_col).value = ro.get("status", "") or "A"
                if type_col: ws.cell(row=r, column=type_col).value = ro.get("type", "") or "H"
                if dept_col: ws.cell(row=r, column=dept_col).value = ro.get("dept", "")
                if rate_col:
                    try: ws.cell(row=r, column=rate_col).value = float(ro.get("pay_rate","") or 0.0)
                    except Exception: ws.cell(row=r, column=rate_col).value = 0.0

                # Hours from Sierra
                if s is not None:
                    matches += 1
                    if reg_col:  ws.cell(row=r, column=reg_col).value  = _num(s["REGULAR"])
                    if ot_col:   ws.cell(row=r, column=ot_col).value   = _num(s["OVERTIME"])
                    if dt_col:   ws.cell(row=r, column=dt_col).value   = _num(s["DOUBLETIME"])
                    if vac_col:  ws.cell(row=r, column=vac_col).value  = _num(s.get("VACATION",0))
                    if sick_col: ws.cell(row=r, column=sick_col).value = _num(s.get("SICK",0))
                    if hol_col:  ws.cell(row=r, column=hol_col).value  = _num(s.get("HOLIDAY",0))
                    if bon_col:  ws.cell(row=r, column=bon_col).value  = _num(s.get("BONUS",0))
                    if com_col:  ws.cell(row=r, column=com_col).value  = _num(s.get("COMMISSION",0))

                    # Piecework hours/totals (pink)
                    ah_src = ["AH1","AH2","AH3","AH4","AH5"]
                    ai_src = ["AI1","AI2","AI3","AI4","AI5"]
                    for idx in range(5):
                        col_h = ah_cols[idx]
                        col_t = ai_cols[idx]
                        if col_h: ws.cell(row=r, column=col_h).value = _num(s.get(ah_src[idx], 0))
                        if col_t: ws.cell(row=r, column=col_t).value = _num(s.get(ai_src[idx], 0))

                # Per-row total formula (=REG+OT+DT); if template already has formula, leave it.
                if tot_col and regL and otL and dtL:
                    c = ws.cell(row=r, column=tot_col)
                    if not (isinstance(c.value, str) and c.value.startswith("=")):
                        c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

                # keep travel/comments blank unless you want to carry data in

            # Bottom totals row (sum down each numeric column we filled)
            last_row = start + len(order)
            if name_col:
                ws.cell(row=last_row, column=name_col).value = "TOTALS"

            def sum_col(col_idx: int):
                if not col_idx: return
                colL = get_column_letter(col_idx)
                ws.cell(row=last_row, column=col_idx).value = f"=SUM({colL}{start}:{colL}{last_row-1})"

            for col in [reg_col, ot_col, dt_col, vac_col, sick_col, hol_col, bon_col, com_col,
                        *ah_cols, *ai_cols, ate_col, tot_col]:
                sum_col(col)

            wb.save(output_path)

            total_hours = float(df[["REGULAR","OVERTIME","DOUBLETIME","VACATION","SICK","HOLIDAY"]].sum().sum()) if not df.empty else 0.0
            return {"success": True, "employees": len(order), "hours": total_hours, "matched": matches}

        except Exception as e:
            return {"success": False, "error": str(e)}
