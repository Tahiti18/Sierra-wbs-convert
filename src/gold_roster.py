# src/gold_roster.py
from __future__ import annotations
import csv
from pathlib import Path
from typing import Dict, List, Optional

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
ORDER_PATH = DATA_DIR / "gold_master_order.txt"
ROSTER_CSV = DATA_DIR / "gold_master_roster.csv"  # Name,SSN,Status,Type,Department,Pay Rate (headers can be in any case)
TEMPLATE_XLSX = DATA_DIR / "wbs_template.xlsx"    # optional but strongly recommended

def load_order() -> List[str]:
    if not ORDER_PATH.exists():
        raise FileNotFoundError(f"gold_master_order.txt not found at {ORDER_PATH}")
    names = [ln.strip() for ln in ORDER_PATH.read_text(encoding="utf-8").splitlines()]
    return [n for n in names if n]

def _norm(h: str) -> str:
    return h.strip().lower().replace(" ", "")

def load_roster_csv() -> Dict[str, Dict[str, str]]:
    """
    Returns a dict keyed by 'Employee Name' -> details dict including SSN, Status, Type, Department, PayRate (if present).
    Header names are matched case-insensitively.
    """
    if not ROSTER_CSV.exists():
        return {}
    with ROSTER_CSV.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        # map headers to canonical keys
        canon_map = {}
        for h in r.fieldnames or []:
            k = _norm(h)
            if k in ("employeename","name"): canon_map[h] = "Employee Name"
            elif k in ("ssn",): canon_map[h] = "SSN"
            elif k in ("status",): canon_map[h] = "Status"
            elif k in ("type",): canon_map[h] = "Type"
            elif k in ("department","dept"): canon_map[h] = "Department"
            elif k in ("payrate","pay rate"): canon_map[h] = "PayRate"
            elif k in ("empid","employeeid"): canon_map[h] = "EmpID"
            else: canon_map[h] = h  # keep anything else
        out: Dict[str, Dict[str, str]] = {}
        for row in r:
            norm = { canon_map.get(k,k): (v or "").strip() for k,v in row.items() }
            name = norm.get("Employee Name","").strip()
            if not name:
                continue
            out[name] = norm
        return out

def ssn_for(name: str, roster: Dict[str, Dict[str, str]]) -> str:
    return (roster.get(name, {}).get("SSN") or "").strip()

def template_path() -> Optional[Path]:
    return TEMPLATE_XLSX if TEMPLATE_XLSX.exists() else None
