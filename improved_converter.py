# improved_converter.py — robust Sierra → WBS converter (final)
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# ---------- repo paths ----------
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

# ---------- small helpers ----------
def _norm(s) -> str:
    return ("" if s is None else str(s)).strip().lower()

def _canon_name(s) -> str:
    s = ("" if s is None else str(s)).strip()
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

def _load_order(path: Path = ORDER_TXT) -> List[str]:
    if not path.exists():
        return []
    return [ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()]

# ---------- template header location ----------
def _find_header_row(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Scan first ~150 rows; map canonical column names to indexes.
    We require an 'Employee Name' cell in the header band.
    """
    cmap: Dict[str, int] = {}
    for r in range(1, min(ws.max_row, 150) + 1):
        labels = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
        joined = "|".join(_norm(v).replace(" ", "") for v in labels)
        if "employeename" in joined:
            # build column map
            for i, raw in enumerate(labels, start=1):
                k = _norm(raw)
                k0 = k.replace(" ", "")
                if k0 in ("employeename", "name"):
                    cmap["Employee Name"] = i
                elif k0 in ("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno"):
                    cmap["SSN"] = i
                elif k0 in ("status",):
                    cmap["Status"] = i
                elif k0 in ("type",):
                    cmap["Type"] = i
                elif k0 in ("dept", "department"):
                    cmap["Dept"] = i
                elif k0 in ("payrate", "pay", "payrate:"):
                    cmap["Pay Rate"] = i
                elif k0 == "regular":
                    cmap["REGULAR"] = i
                elif k0 in ("overtime", "ot"):
                    cmap["OVERTIME"] = i
                elif k0 in ("doubletime", "doubletime", "double time"):
                    cmap["DOUBLETIME"] = i
                elif k0 in ("totals", "total", "sum"):
                    cmap["Totals"] = i
            if "Employee Name" in cmap:
                return r, cmap
    raise ValueError("Template scan failed: could not locate the header row containing 'Employee Name'.")

def _first_data_row(hdr_row: int) -> int:
    return hdr_row + 1

# ---------- converter ----------
class SierraToWBSConverter:
    """
    Robust Sierra → WBS converter.
    - Accepts WEEKLY sheet or first sheet.
    - Detects header row at row 8 (classic) or scans row 1 if needed.
    - Tolerates column aliasing (REGULAR/OVERTIME/DOUBLETIME variants).
    - Fills roster fields from gold_master_roster.csv.
    - Locks output order to gold_master_order.txt.
    - Preserves template formatting and writes the pink 'Totals' formula if missing.
    """

    def __init__(self, gold_master_order_path: str | None = None):
        # for /health logging in main.py
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        self.gold_master_order: List[str] = _load_order(path)

    # ---------- used by /validate and /process ----------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Returns DataFrame with: Name, REGULAR, OVERTIME, DOUBLETIME, Hours, __canon
        Never raises on shape — produces empty DF if truly unreadable.
        """
        # Try WEEKLY, header at row 8 (0-indexed 7)
        def _try_read(sheet, header):
            try:
                return pd.read_excel(input_path, sheet_name=sheet, header=header).dropna(how="all")
            except Exception:
                return None

        df = _try_read("WEEKLY", 7)
        if df is None or "Employee Name" not in df.columns:
            # try first sheet, same header row
            df = _try_read(0, 7)
        if df is None:
            # last resort: first sheet, header row 0
            df = _try_read(0, 0)
        if df is None:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        # Some exports repeat the header in the first data row — drop it
        if not df.empty and "Employee Name" in df.columns and str(df.iloc[0].get("Employee Name", "")).strip() == "Employee Name":
            df = df.iloc[1:]

        # pick the Name column
        name_col = None
        for cand in ["Employee Name", "Name", "Unnamed: 2"]:
            if cand in df.columns:
                name_col = cand
                break
        if not name_col:
            # scan any column that looks like names
            for c in df.columns:
                if re.search(r"name", str(c), re.I):
                    name_col = c
                    break
        if not name_col:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})
        out = out[out["Name"].str.len() > 0]

        def to_num(series_like):
            return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

        # map possible hour columns
        alias = {
            "REGULAR":   ["REGULAR", "Regular", "A01"],
            "OVERTIME":  ["OVERTIME", "Overtime", "OT", "A02"],
            "DOUBLETIME":["DOUBLETIME", "Double Time", "DOUBLE TIME", "A03"],
        }
        for dst, srcs in alias.items():
            col = next((c for c in srcs if c in df.columns), None)
            out[dst] = to_num(df[col]) if col else 0.0

        out["Hours"] = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out["__canon"] = out["Name"].map(_canon_name)

        # strip obviously non-employee rows
        bad = ("signature", "certify", "gross", "week of", "prepared by")
        out = out[~out["Name"].str.lower().str.contains("|".join(bad), regex=True, na=False)]
        out.reset_index(drop=True, inplace=True)
        return out

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        """
        Read gold_master_roster.csv → canonical name map with ssn/status/type/department/pay_rate
        If missing, return {} (converter will still run using hours only).
        """
        if not ROSTER_CSV.exists():
            print(f"[WARN] Roster not found: {ROSTER_CSV}")
            return {}

        try:
            df = pd.read_csv(ROSTER_CSV)
        except Exception as e:
            print(f"[WARN] Failed to read roster CSV: {e}")
            return {}

        # normalize header names (accept loose variations)
        header_map = {c: _norm(c).replace(" ", "") for c in df.columns}

        def pick(*aliases):
            al = [a.replace(" ", "").lower() for a in aliases]
            for c, key in header_map.items():
                if key in al:
                    return c
            return None

        name_col = pick("employeename", "name")
        ssn_col  = pick("ssn", "socialsecuritynumber", "socialsecurity#", "socialsecurityno")
        stat_col = pick("status")
        type_col = pick("type")
        dept_col = pick("dept", "department")
        rate_col = pick("payrate", "payrate:", "pay", "pay rate")

        if not name_col:
            print("[WARN] Roster missing an 'Employee Name' column.")
            return {}

        out: Dict[str, Dict[str, str]] = {}
        for _, r in df.iterrows():
            nm = str(r.get(name_col, "")).strip()
            if not nm:
                continue
            key = _canon_name(nm)
            out[key] = {
                "ssn": "" if ssn_col is None or pd.isna(r.get(ssn_col)) else str(r.get(ssn_col)).strip(),
                "status": "" if stat_col is None or pd.isna(r.get(stat_col)) else str(r.get(stat_col)).strip(),
                "type": "" if type_col is None or pd.isna(r.get(type_col)) else str(r.get(type_col)).strip(),
                "dept": "" if dept_col is None or pd.isna(r.get(dept_col)) else str(r.get(dept_col)).strip(),
                "pay_rate": "" if rate_col is None or pd.isna(r.get(rate_col)) else str(r.get(rate_col)).strip(),
            }

        if not out:
            print("[WARN] Roster parsed but produced 0 rows after normalization.")
        return out

    # ---------- main entry ----------
    def convert(self, input_path: str, output_path: str) -> Dict:
        """
        Fill the template with roster + hours and save to output_path.
        """
        try:
            order = self.gold_master_order[:] or _load_order()
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

            hdr_row, cmap = _find_header_row(ws)
            start = _first_data_row(hdr_row)

            name_col = cmap["Employee Name"]
            ssn_col  = cmap.get("SSN")
            stat_col = cmap.get("Status")
            type_col = cmap.get("Type")
            dept_col = cmap.get("Dept")
            rate_col = cmap.get("Pay Rate")
            reg_col  = cmap.get("REGULAR")
            ot_col   = cmap.get("OVERTIME")
            dt_col   = cmap.get("DOUBLETIME")
            tot_col  = cmap.get("Totals")

            regL = get_column_letter(reg_col) if reg_col else None
            otL  = get_column_letter(ot_col)  if ot_col  else None
            dtL  = get_column_letter(dt_col)  if dt_col  else None

            matched = 0
            for i, emp in enumerate(order):
                r = start + i
                ws.cell(row=r, column=name_col).value = emp

                key = _canon_name(emp)
                s   = sierra_map.get(key)
                ro  = roster.get(key, {})

                # roster fields
                if ssn_col:  ws.cell(row=r, column=ssn_col).value  = ro.get("ssn", "")
                if stat_col: ws.cell(row=r, column=stat_col).value = ro.get("status", "") or "A"
                if type_col: ws.cell(row=r, column=type_col).value = ro.get("type", "") or "H"
                if dept_col: ws.cell(row=r, column=dept_col).value = ro.get("dept", "")

                if rate_col:
                    try:
                        rate_val = float(ro.get("pay_rate", "") or 0.0)
                    except Exception:
                        rate_val = 0.0
                    ws.cell(row=r, column=rate_col).value = rate_val

                # hours
                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0
                if reg_col: ws.cell(row=r, column=reg_col).value = reg
                if ot_col:  ws.cell(row=r, column=ot_col).value  = ot
                if dt_col:  ws.cell(row=r, column=dt_col).value  = dt
                if s is not None:
                    matched += 1

                # ensure pink totals (only if template cell lacks a formula already)
                if tot_col:
                    c = ws.cell(row=r, column=tot_col)
                    if not (isinstance(c.value, str) and c.value.startswith("=")):
                        if regL and otL and dtL:
                            c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

            wb.save(output_path)

            total_hours = float(sierra_df[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum().sum()) if not sierra_df.empty else 0.0
            if matched == 0:
                print("[WARN] No Sierra rows matched names in gold order after canonicalization.")

            return {"success": True, "employees": len(order), "total_hours": total_hours}

        except Exception as e:
            return {"success": False, "error": str(e)}
