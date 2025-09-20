# improved_converter.py  — debug-hardened Sierra → WBS converter
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import io
import re
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ---------- paths ----------
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ORDER_TXT = DATA / "gold_master_order.txt"
ROSTER_CSV = DATA / "gold_master_roster.csv"
TEMPLATE_XLSX = DATA / "wbs_template.xlsx"
TARGET_SHEET = "WEEKLY"

# ---------- small utils ----------
def _norm(s) -> str:
    return ("" if s is None else str(s)).strip()

def _canon_name(s: str) -> str:
    s = _norm(s)
    s = s.replace(".", "")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s*,\s*", ", ", s)
    s = s.replace(" ,", ",")
    return s.lower()

def _to_num(v) -> float:
    try:
        if v is None or (isinstance(v, str) and v.strip() == ""):
            return 0.0
        return float(v)
    except Exception:
        return 0.0

def _first_existing(*candidates: str, in_cols: List[str]) -> Optional[str]:
    for c in candidates:
        if c in in_cols:
            return c
    return None

# ---------- header scan ----------
def find_header_row_and_map(xls_path: str) -> Tuple[int, Dict[str, str], pd.DataFrame]:
    """
    Scan first sheet for a row that contains the Sierra headers.
    Returns: (header_row_index, column_map, dataframe_read_with_that_header)
    column_map keys: 'Name', 'REG', 'OT', 'DT'
    """
    # Try 'WEEKLY' first (many of your golds use it), else first sheet
    for sheet_name in ("WEEKLY", 0):
        try:
            raw = pd.read_excel(xls_path, sheet_name=sheet_name, header=None, dtype=str)
        except Exception:
            continue
        if raw.empty:
            continue

        # Look at first ~40 rows to find the header
        best = None
        for r in range(min(40, len(raw))):
            row = [(_norm(v)).lower().replace(" ", "") for v in raw.iloc[r].tolist()]
            if not row:
                continue

            has_name = any(k in row for k in ("employeename", "name"))
            has_any_hours = any(k in row for k in ("regular", "overtime", "ot", "doubletime", "doubletime"))
            if has_name and has_any_hours:
                best = r
                break

        if best is None:
            # fall back to row 7 (index 7) because many Sierra exports start there
            hdr = 7 if raw.shape[0] > 8 else 0
        else:
            hdr = best

        df = pd.read_excel(xls_path, sheet_name=sheet_name, header=hdr)
        df.columns = [str(c).strip() for c in df.columns]
        cols = list(df.columns)

        # Map columns with aliases
        name_col = _first_existing("Employee Name", "Name", "EMPLOYEE NAME", "Unnamed: 2", in_cols=cols)
        reg_col  = _first_existing("REGULAR", "Regular", "REG", in_cols=cols)
        ot_col   = _first_existing("OVERTIME", "Overtime", "OT", in_cols=cols)
        dt_col   = _first_existing("DOUBLETIME", "Double Time", "DOUBLE TIME", "DT", in_cols=cols)

        colmap = {"Name": name_col, "REG": reg_col, "OT": ot_col, "DT": dt_col}
        return hdr, colmap, df

    raise ValueError("Could not open Sierra Excel (no readable sheets).")

# ---------- main converter ----------
class SierraToWBSConverter:
    def __init__(self, order_path: Optional[str] = None):
        p = Path(order_path) if order_path else ORDER_TXT
        self.gold_master_order: List[str] = []
        if p.exists():
            self.gold_master_order = [ln.strip() for ln in p.read_text(encoding="utf-8").splitlines() if ln.strip()]

    # --- DEBUG: parse only, with lots of prints
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        print(f"[DEBUG] parse_sierra_file: {input_path}", flush=True)

        hdr_row, cmap, df = find_header_row_and_map(input_path)
        print(f"[DEBUG] Header detected at Excel row index = {hdr_row}", flush=True)
        print(f"[DEBUG] Columns found: {list(df.columns)}", flush=True)
        print(f"[DEBUG] Column map => Name: {cmap.get('Name')}, REG: {cmap.get('REG')}, "
              f"OT: {cmap.get('OT')}, DT: {cmap.get('DT')}", flush=True)

        if not cmap.get("Name"):
            raise ValueError("Could not find 'Employee Name' column in Sierra file.")

        # Build a clean frame with numeric hours
        out = pd.DataFrame({
            "Name": df[cmap["Name"]].astype(str).str.strip()
        })

        def pick(col_key: str) -> pd.Series:
            col = cmap.get(col_key)
            if not col or col not in df.columns:
                return pd.Series([0.0] * len(out))
            return pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        out["REGULAR"]   = pick("REG")
        out["OVERTIME"]  = pick("OT")
        out["DOUBLETIME"]= pick("DT")
        out["Hours"]     = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out["__canon"]   = out["Name"].map(_canon_name)

        # Filter garbage rows
        before = len(out)
        out = out[out["Name"].astype(str).str.strip() != ""]
        # Drop header echo rows, totals, signatures, etc.
        mask_bad = out["Name"].str.lower().str.contains(
            r"(employee\s*name|total|totals|signature|certify|week of|gross|grand)", regex=True, na=False
        )
        out = out[~mask_bad]
        after = len(out)

        print(f"[DEBUG] Rows before cleanup={before}, after cleanup={after}", flush=True)
        print(f"[DEBUG] Total hours parsed = {float(out['Hours'].sum()):.3f}", flush=True)
        return out.reset_index(drop=True)

    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        out: Dict[str, Dict[str, str]] = {}
        if not ROSTER_CSV.exists():
            print(f"[WARN] Roster not found: {ROSTER_CSV}", flush=True)
            return out

        try:
            r = pd.read_csv(ROSTER_CSV, dtype=str).fillna("")
        except Exception as e:
            print(f"[WARN] Failed to read roster CSV: {e}", flush=True)
            return out

        def pick(*aliases) -> Optional[str]:
            aliases = [a.replace(" ", "").lower() for a in aliases]
            for c in r.columns:
                if c.replace(" ", "").lower() in aliases:
                    return c
            return None

        name_c = pick("employeename", "name")
        ssn_c  = pick("ssn", "socialsecuritynumber", "socialsecurity#")
        st_c   = pick("status")
        tp_c   = pick("type")
        dp_c   = pick("dept", "department")
        pr_c   = pick("payrate", "pay rate", "rate")

        if not name_c:
            print("[WARN] Roster missing Employee Name column.", flush=True)
            return out

        for _, row in r.iterrows():
            nm = _norm(row.get(name_c, ""))
            if not nm:
                continue
            k = _canon_name(nm)
            out[k] = {
                "ssn": _norm(row.get(ssn_c, "")),
                "status": _norm(row.get(st_c, "")) or "A",
                "type": _norm(row.get(tp_c, "")) or "H",
                "dept": _norm(row.get(dp_c, "")),
                "pay_rate": _norm(row.get(pr_c, "")),
            }
        print(f"[DEBUG] Roster rows loaded = {len(out)}", flush=True)
        return out

    def convert(self, input_path: str, output_path: str) -> Dict:
        # Parse Sierra rows
        df = self.parse_sierra_file(input_path)
        sierra_map = {row["__canon"]: row for _, row in df.iterrows()}

        # Load gold order
        order = self.gold_master_order[:]
        if not order and ORDER_TXT.exists():
            order = [ln.strip() for ln in ORDER_TXT.read_text(encoding="utf-8").splitlines() if ln.strip()]
        if not order:
            raise ValueError("gold_master_order.txt missing/empty")

        # Load roster attributes
        roster = self._load_roster()

        # Load template
        if not TEMPLATE_XLSX.exists():
            raise ValueError(f"Template not found: {TEMPLATE_XLSX}")
        wb = load_workbook(TEMPLATE_XLSX, data_only=False)
        if TARGET_SHEET not in wb.sheetnames:
            raise ValueError(f"Template missing sheet '{TARGET_SHEET}'")
        ws: Worksheet = wb[TARGET_SHEET]

        # Locate WBS header row/columns inside template
        hdr_row = None
        col_map: Dict[str, int] = {}
        for r in range(1, min(ws.max_row, 150) + 1):
            row_vals = [(_norm(c.value)).lower().replace(" ", "") for c in ws[r]]
            if "employeename" in row_vals and "ssn" in row_vals:
                hdr_row = r
                # build map
                for j, c in enumerate(ws[r], start=1):
                    k = (_norm(c.value)).lower().replace(" ", "")
                    if k == "employeename": col_map["name"] = j
                    elif k == "ssn": col_map["ssn"] = j
                    elif k == "regular": col_map["reg"] = j
                    elif k in ("overtime", "ot"): col_map["ot"] = j
                    elif k in ("doubletime", "doubletime"): col_map["dt"] = j
                    elif k == "status": col_map["status"] = j
                    elif k == "type": col_map["type"] = j
                    elif k in ("payrate", "payrate", "pay"): col_map["rate"] = j
                    elif k in ("dept", "department"): col_map["dept"] = j
                    elif k in ("totals", "total", "sum"): col_map["tot"] = j
                break
        if hdr_row is None:
            raise ValueError("Could not find header row in template.")

        start_row = hdr_row + 1
        regL = get_column_letter(col_map.get("reg", 0)) if "reg" in col_map else None
        otL  = get_column_letter(col_map.get("ot", 0))  if "ot" in col_map else None
        dtL  = get_column_letter(col_map.get("dt", 0))  if "dt" in col_map else None
        totC = col_map.get("tot")

        matches = 0
        for i, emp in enumerate(order):
            r = start_row + i
            k = _canon_name(emp)
            s = sierra_map.get(k, None)
            ro = roster.get(k, {})

            # Name & SSN
            if "name" in col_map: ws.cell(row=r, column=col_map["name"]).value = emp
            if "ssn"  in col_map: ws.cell(row=r, column=col_map["ssn"]).value  = ro.get("ssn", "")

            # personnel fields
            if "status" in col_map: ws.cell(row=r, column=col_map["status"]).value = ro.get("status", "A")
            if "type"   in col_map: ws.cell(row=r, column=col_map["type"]).value   = ro.get("type", "H")
            if "dept"   in col_map: ws.cell(row=r, column=col_map["dept"]).value   = ro.get("dept", "")

            if "rate" in col_map:
                try:
                    ws.cell(row=r, column=col_map["rate"]).value = float(ro.get("pay_rate", "") or 0.0)
                except Exception:
                    ws.cell(row=r, column=col_map["rate"]).value = 0.0

            # hours
            reg = _to_num(s["REGULAR"]) if s is not None else 0.0
            ot  = _to_num(s["OVERTIME"]) if s is not None else 0.0
            dt  = _to_num(s["DOUBLETIME"]) if s is not None else 0.0
            if "reg" in col_map: ws.cell(row=r, column=col_map["reg"]).value = reg
            if "ot"  in col_map: ws.cell(row=r, column=col_map["ot"]).value  = ot
            if "dt"  in col_map: ws.cell(row=r, column=col_map["dt"]).value  = dt
            if s is not None:
                matches += 1

            # totals column: ensure =REG+OT+DT if template has no formula there
            if totC:
                c = ws.cell(row=r, column=totC)
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")) and regL and otL and dtL:
                    c.value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

        print(f"[DEBUG] Employees in gold order: {len(order)}; matched Sierra rows: {matches}", flush=True)
        wb.save(output_path)

        return {
            "success": True,
            "employees": len(order),
            "matched": matches,
            "hours": float(df["Hours"].sum()) if not df.empty else 0.0
        }
