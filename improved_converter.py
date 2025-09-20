# improved_converter.py — template-locked writer + robust Sierra parser
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

def _load_order(path: Optional[Path] = None) -> List[str]:
    p = path or ORDER_TXT
    if not p.exists():
        return []
    return [ln.strip() for ln in p.read_text(encoding="utf-8").splitlines() if ln.strip()]

def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Find the header row that contains 'Employee Name' and build a column map.
    Keeps your template’s headings & the ‘code line’ (A01/A02/… AH1/AI1 …).
    """
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 200) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        low = [_norm(v).replace(" ", "") for v in labels]
        if "employeename" in low:
            # Build map for the row with spelled-out headings
            for i, v in enumerate(labels, start=1):
                k = _norm(v)
                k0 = k.replace(" ", "")
                if k0 in ("employeename",):
                    cmap["Employee Name"] = i
                elif k0 in ("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno"):
                    cmap["SSN"] = i
                elif k0 == "status":
                    cmap["Status"] = i
                elif k0 == "type":
                    cmap["Type"] = i
                elif k0 in ("payrate", "pay", "payrate:"):
                    cmap["Pay Rate"] = i
                elif k0 in ("dept", "department"):
                    cmap["Dept"] = i
                elif k0 == "regular":
                    cmap["REGULAR"] = i
                elif k0 in ("overtime", "ot"):
                    cmap["OVERTIME"] = i
                elif k0 in ("doubletime", "double time"):
                    cmap["DOUBLETIME"] = i
                elif k0 in ("bonus",):
                    cmap["BONUS"] = i
                elif k0 in ("commission",):
                    cmap["COMMISSION"] = i
                elif k0 in ("vacation",):
                    cmap["VACATION"] = i
                elif k0 in ("sick",):
                    cmap["SICK"] = i
                elif k0 in ("holiday",):
                    cmap["HOLIDAY"] = i
                elif k0 in ("totals", "total", "sum"):
                    cmap["Totals"] = i
            return r, cmap
    raise ValueError("Template header row containing 'Employee Name' was not found.")

def _first_data_row(h: int) -> int:
    return h + 1

# ---------------- converter ----------------
class SierraToWBSConverter:
    """
    Template-locked converter.
    Writes ONLY the body rows under your template’s headers,
    preserving your exact formatting, the “# V / # U …” banner rows, the
    spelled-out headings row AND the A01/A02/…/AH1/AI1 code row.
    """
    def __init__(self, gold_master_order_path: str | None = None):
        self.gold_master_order: List[str] = []
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        if path.exists():
            self.gold_master_order = _load_order(path)

    # ---------- used by /validate ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Return DF with Name, REGULAR, OVERTIME, DOUBLETIME, Hours, __canon.

        Robust to Sierra files where WEEKLY starts with a banner and real headers
        are on row 8, or when the first sheet is used, or when columns are
        labeled slightly differently (Regular vs REGULAR etc.).
        """
        # Try WEEKLY with header at row 8 (0-indexed 7), then fallbacks
        tried = []
        for sheet, header in [("WEEKLY", 7), (0, 7), (0, 0)]:
            try:
                df = pd.read_excel(input_path, sheet_name=sheet, header=header).dropna(how="all")
                tried.append((sheet, header, True))
                break
            except Exception:
                tried.append((sheet, header, False))
                df = None
        if df is None or df.empty:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        # remove redundant header row if Sierra duplicated the headings
        if "Employee Name" in df.columns and str(df.iloc[0].get("Employee Name", "")).strip() == "Employee Name":
            df = df.iloc[1:].copy()

        # pick a name column
        name_col = next((c for c in ["Employee Name", "Name", "Unnamed: 2"] if c in df.columns), None)
        if not name_col:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})

        def to_num(series_like):
            return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

        def pick(df_, aliases: List[str]):
            for a in aliases:
                if a in df_.columns:
                    return a
            return None

        reg_c = pick(df, ["REGULAR", "Regular", "Reg"])
        ot_c  = pick(df, ["OVERTIME", "Overtime", "OT"])
        dt_c  = pick(df, ["DOUBLETIME", "Double Time", "DOUBLE TIME", "DT"])

        out["REGULAR"]   = to_num(df[reg_c]) if reg_c else 0.0
        out["OVERTIME"]  = to_num(df[ot_c]) if ot_c else 0.0
        out["DOUBLETIME"]= to_num(df[dt_c]) if dt_c else 0.0
        out["Hours"]     = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)

        # keep only real names
        out = out[out["Name"].astype(str).str.strip().str.len() > 0].copy()
        out["__canon"] = out["Name"].map(_canon_name)
        return out.reset_index(drop=True)

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        """
        Return canonical_name -> {ssn,status,type,dept,pay_rate}
        """
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
            aliases = [a.replace(" ", "").lower() for a in aliases]
            for c, k in hdr.items():
                if k in aliases:
                    return c
            return None

        name_col = col("employeename", "name")
        ssn_col  = col("ssn","socialsecuritynumber","socialsecurity#","socialsecurityno")
        stat_col = col("status")
        type_col = col("type")
        dept_col = col("dept","department")
        rate_col = col("payrate","payrate:","pay")

        if not name_col:
            print("[WARN] Roster missing Employee Name column.")
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
        """
        1) Open data/wbs_template.xlsx (sheet WEEKLY) and locate the spelled-out header row.
        2) Write body rows strictly in gold_master_order.txt order.
        3) Fill SSN/Status/Type/Dept/Pay Rate from roster, REG/OT/DT from Sierra.
        4) Ensure Totals column has a =REG+OT+DT formula when missing.
        5) Leave pink columns (PC HRS/TTL …) in place; write zeros if template
           cells are empty so the area is never blank.
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

            # header map
            h, cmap = _find_header_row(ws)
            start = _first_data_row(h)

            # critical columns (by header names from template)
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

            # letters for totals formula
            regL = get_column_letter(reg_col) if reg_col else None
            otL  = get_column_letter(ot_col) if ot_col else None
            dtL  = get_column_letter(dt_col) if dt_col else None

            # PRE-FILL the pink “PC … / TTL …” region with zeros so it’s never blank.
            # Find the code row immediately BELOW your spelled-out header (e.g. AH1, AI1, …)
            code_row_idx = h + 1
            code_row = [c.value for c in ws[code_row_idx]]
            code_labels = [str(v).strip() if v is not None else "" for v in code_row]
            # Identify all “AH1/AI1 … AH5/AI5” + “ATE/Comments/Totals” columns by position
            pink_cols = []
            for idx, label in enumerate(code_labels, start=1):
                lab = label.upper().replace(" ", "")
                if lab in {"AH1","AI1","AH2","AI2","AH3","AI3","AH4","AI4","AH5","AI5","ATE"}:
                    pink_cols.append(idx)

            # CLEAR existing data area (optional safety; comment out if template already clean)
            max_rows_to_write = max(len(order), 120)  # safety bound
            for r in range(start, start + max_rows_to_write):
                # write zeros in pink cols if there is a cell and it’s empty
                for cidx in pink_cols:
                    c = ws.cell(row=r, column=cidx)
                    if c.value in (None, ""):
                        c.value = 0

            matches = 0
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
                        val = float(ro.get("pay_rate", "") or 0.0)
                    except Exception:
                        val = 0.0
                    ws.cell(row=r, column=rate_col).value = val

                # Hours from Sierra (REG/OT/DT)
                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0
                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt

                # zero out vac/sick/holiday/bonus/commission if template cell is empty
                for cidx in [vac_col, sick_col, hol_col, bon_col, com_col]:
                    if cidx:
                        cell = ws.cell(row=r, column=cidx)
                        if cell.value in (None, ""):
                            cell.value = 0

                # Row total: if template cell lacks a formula, write =REG+OT+DT
                if tot_col:
                    c = ws.cell(row=r, column=tot_col)
                    has_formula = isinstance(c.value, str) and c.value.startswith("=")
                    if not has_formula and regL and otL and dtL:
                        c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

                if s is not None:
                    matches += 1

            wb.save(output_path)

            total_hours = float(sierra_df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not sierra_df.empty else 0.0
            if matches == 0:
                print("[WARN] No Sierra rows matched names in gold order after canonicalization.")

            return {"success": True, "employees": len(order), "hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
