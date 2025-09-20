# improved_converter.py — stable Sierra→WBS using template; fixes totals & headers
from __future__ import annotations
from pathlib import Path
from typing import Dict, List
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

def _canon_name(s: str) -> str:
    s = "" if s is None else str(s)
    s = re.sub(r"\s+", " ", s.strip())
    s = s.replace(".", "")
    s = re.sub(r"\s*,\s*", ", ", s)
    return s.lower()

def _num(v) -> float:
    try:
        if v is None or (isinstance(v, str) and v.strip() == ""):
            return 0.0
        return float(v)
    except Exception:
        return 0.0

class SierraToWBSConverter:
    def __init__(self, gold_master_order_path: str | None = None):
        path = Path(gold_master_order_path) if gold_master_order_path else ORDER_TXT
        self.gold_master_order: List[str] = []
        if path.exists():
            self.gold_master_order = [ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()]

    # ---------------- parse Sierra ----------------
    def parse_sierra_file(self, input_path: str) -> pd.DataFrame:
        """
        Returns DF with columns: Name, REGULAR, OVERTIME, DOUBLETIME, Hours, __canon.
        Works with WEEKLY sheet (header row 8) or falls back to first sheet.
        """
        def _read(try_weekly=True):
            if try_weekly:
                return pd.read_excel(input_path, sheet_name="WEEKLY", header=7).dropna(how="all")
            return pd.read_excel(input_path, sheet_name=0, header=7).dropna(how="all")

        try:
            df = _read(True)
        except Exception:
            df = _read(False)

        # Occasionally an extra header row is duplicated—skip it
        if not df.empty and "Employee Name" in df.columns and str(df.iloc[0]["Employee Name"]).strip() == "Employee Name":
            df = df.iloc[1:]

        # pick name column
        name_col = None
        for c in ["Employee Name", "Name", "Unnamed: 2"]:
            if c in df.columns:
                name_col = c
                break
        if not name_col:
            return pd.DataFrame(columns=["Name", "REGULAR", "OVERTIME", "DOUBLETIME", "Hours", "__canon"])

        out = pd.DataFrame({"Name": df[name_col].astype(str).str.strip()})

        # map hours columns with aliases
        def to_num(series_like):
            return pd.to_numeric(series_like, errors="coerce").fillna(0.0)

        aliases = {
            "REGULAR":   ["A01", "REGULAR", "Regular"],
            "OVERTIME":  ["A02", "OVERTIME", "Overtime", "OT"],
            "DOUBLETIME":["A03", "DOUBLETIME", "Double Time", "DOUBLE TIME"]
        }
        for dst, srcs in aliases.items():
            col = next((c for c in srcs if c in df.columns), None)
            out[dst] = to_num(df[col]) if col else 0.0

        out["Hours"] = out[["REGULAR", "OVERTIME", "DOUBLETIME"]].sum(axis=1)
        out = out[out["Name"].str.strip().ne("")]
        out["__canon"] = out["Name"].map(_canon_name)
        return out.reset_index(drop=True)

    # ---------------- roster ----------------
    def _load_roster(self) -> Dict[str, Dict[str, str]]:
        if not ROSTER_CSV.exists():
            print(f"[WARN] Roster not found: {ROSTER_CSV}")
            return {}
        try:
            df = pd.read_csv(ROSTER_CSV)
        except Exception as e:
            print(f"[WARN] Failed to read roster CSV: {e}")
            return {}

        cols = {c.lower().replace(" ", ""): c for c in df.columns}
        def pick(*keys):
            for k in keys:
                k = k.replace(" ", "").lower()
                if k in cols: return cols[k]
            return None

        name_c = pick("employeename", "name")
        ssn_c  = pick("ssn","socialsecuritynumber","socialsecurity#")
        stat_c = pick("status")
        type_c = pick("type")
        dept_c = pick("dept","department")
        rate_c = pick("payrate","pay rate","rate")
        if not name_c:
            return {}

        out: Dict[str, Dict[str, str]] = {}
        for _, r in df.iterrows():
            nm = str(r.get(name_c, "")).strip()
            if not nm: continue
            key = _canon_name(nm)
            out[key] = {
                "ssn":     "" if ssn_c  is None or pd.isna(r.get(ssn_c))  else str(r.get(ssn_c)).strip(),
                "status":  "" if stat_c is None or pd.isna(r.get(stat_c)) else str(r.get(stat_c)).strip(),
                "type":    "" if type_c is None or pd.isna(r.get(type_c)) else str(r.get(type_c)).strip(),
                "dept":    "" if dept_c is None or pd.isna(r.get(dept_c)) else str(r.get(dept_c)).strip(),
                "payrate": "" if rate_c is None or pd.isna(r.get(rate_c)) else str(r.get(rate_c)).strip(),
            }
        return out

    # ---------------- convert using template ----------------
    def convert(self, input_path: str, output_path: str) -> Dict:
        try:
            order = self.gold_master_order[:]
            if not order:
                return {"success": False, "error": "gold_master_order.txt missing/empty"}

            sierra = self.parse_sierra_file(input_path)
            s_map = {row["__canon"]: row for _, row in sierra.iterrows()}
            roster = self._load_roster()

            if not TEMPLATE_XLSX.exists():
                return {"success": False, "error": f"Template not found: {TEMPLATE_XLSX}"}
            wb: Workbook = load_workbook(TEMPLATE_XLSX, data_only=False)
            if TARGET_SHEET not in wb.sheetnames:
                return {"success": False, "error": f"Template missing '{TARGET_SHEET}'"}
            ws: Worksheet = wb[TARGET_SHEET]

            # find header row & exact columns by header text
            header_row = None
            col_idx = {}
            for r in range(1, min(ws.max_row, 150)+1):
                row_vals = [str(c.value).strip() if c.value is not None else "" for c in ws[r]]
                compact = [v.lower().replace(" ", "") for v in row_vals]
                if "employeename" in compact and "ssn" in compact:
                    header_row = r
                    for i, v in enumerate(row_vals, start=1):
                        key = v.lower().replace(" ", "")
                        if key == "ssn": col_idx["SSN"] = i
                        if key == "employeename": col_idx["Name"] = i
                        if key == "status": col_idx["Status"] = i
                        if key == "type": col_idx["Type"] = i
                        if key in ("payrate","pay"): col_idx["PayRate"] = i
                        if key in ("dept","department"): col_idx["Dept"] = i
                        if key == "regular": col_idx["A01"] = i
                        if key == "overtime": col_idx["A02"] = i
                        if key == "doubletime": col_idx["A03"] = i
                        if key == "vacation": col_idx["A06"] = i
                        if key == "sick": col_idx["A07"] = i
                        if key == "holiday": col_idx["A08"] = i
                        if key == "bonus": col_idx["A04"] = i
                        if key == "commission": col_idx["A05"] = i
                        if key in ("totals","total"): col_idx["Totals"] = i
                    break
            if header_row is None:
                return {"success": False, "error": "Could not locate header row in template"}

            start_row = header_row + 1

            # letters for totals formula
            regL = get_column_letter(col_idx["A01"])
            otL  = get_column_letter(col_idx["A02"])
            dtL  = get_column_letter(col_idx["A03"])
            totC = col_idx.get("Totals")

            matches = 0
            for i, emp in enumerate(order):
                r = start_row + i
                key = _canon_name(emp)
                s = s_map.get(key, None)
                ro = roster.get(key, {})

                # Names + SSN / meta
                ws.cell(row=r, column=col_idx["Name"]).value = emp
                if "SSN"     in col_idx: ws.cell(row=r, column=col_idx["SSN"]).value     = ro.get("ssn","")
                if "Status"  in col_idx: ws.cell(row=r, column=col_idx["Status"]).value  = (ro.get("status") or "A")
                if "Type"    in col_idx: ws.cell(row=r, column=col_idx["Type"]).value    = (ro.get("type") or "H")
                if "Dept"    in col_idx: ws.cell(row=r, column=col_idx["Dept"]).value    = ro.get("dept","")
                if "PayRate" in col_idx:
                    try:
                        ws.cell(row=r, column=col_idx["PayRate"]).value = float(ro.get("payrate","") or 0.0)
                    except Exception:
                        ws.cell(row=r, column=col_idx["PayRate"]).value = 0.0

                # Hours
                reg = _num(s["REGULAR"]) if s is not None else 0.0
                ot  = _num(s["OVERTIME"]) if s is not None else 0.0
                dt  = _num(s["DOUBLETIME"]) if s is not None else 0.0
                ws.cell(row=r, column=col_idx["A01"]).value = reg
                ws.cell(row=r, column=col_idx["A02"]).value = ot
                ws.cell(row=r, column=col_idx["A03"]).value = dt

                # leave A04..A08, AH1..AI5, ATE, Comments as-is (template defaults)
                if s is not None:
                    matches += 1

                # Totals column formula (pink)
                if totC:
                    ws.cell(row=r, column=totC).value = f"={regL}{r}+{otL}{r}+{dtL}{r}"

            wb.save(output_path)

            total_hours = float(sierra[["REGULAR","OVERTIME","DOUBLETIME"]].sum().sum()) if not sierra.empty else 0.0
            if matches == 0:
                print("[WARN] no names matched gold order after canonicalization")
            return {"success": True, "employees": len(order), "hours": total_hours}
        except Exception as e:
            return {"success": False, "error": str(e)}
