# src/roster_enforcer.py
import os
from pathlib import Path
import pandas as pd

def _read_roster_order(gold_master_path: str) -> list[str]:
    """
    Reads data/gold_master_order.txt (one name per line) and returns the ordered list.
    """
    if not gold_master_path or not Path(gold_master_path).exists():
        raise FileNotFoundError(f"Gold master order file not found: {gold_master_path}")
    with open(gold_master_path, "r", encoding="utf-8") as f:
        names = [line.strip() for line in f.readlines()]
    # Drop blanks, keep order
    return [n for n in names if n]

def _read_sierra_for_ssn(input_sierra_path: str) -> dict:
    """
    Reads the uploaded Sierra Excel and builds a mapping: name -> ssn.
    Sierra weekly layout: header row at index 7; SSN in 'Unnamed: 1', Name in 'Unnamed: 2'.
    """
    ssn_map = {}
    try:
        df = pd.read_excel(input_sierra_path, sheet_name="WEEKLY", header=7)
        df = df.dropna(how="all")
        # First data row still repeats the header text; skip it
        if (df.iloc[0:1].astype(str).apply(lambda x: (x == 'Employee Name').any(), axis=1).any()):
            df = df.iloc[1:]
        name_col = "Unnamed: 2"
        ssn_col = "Unnamed: 1"
        if name_col not in df.columns or ssn_col not in df.columns:
            return ssn_map
        for _, r in df[[name_col, ssn_col]].dropna(subset=[name_col]).iterrows():
            name = str(r[name_col]).strip()
            ssn = "" if pd.isna(r[ssn_col]) else str(r[ssn_col]).strip()
            if name:
                ssn_map[name] = ssn
    except Exception:
        # Best-effort; if parsing fails we just return empty map
        pass
    return ssn_map

def enforce_roster(output_wbs_path: str, input_sierra_path: str, gold_master_path: str) -> None:
    """
    Opens the just-generated WBS Excel, reorders rows to EXACT gold roster order,
    and inserts missing employees as zero rows. Preserves all columns.
    """
    roster = _read_roster_order(gold_master_path)
    ssn_map = _read_sierra_for_ssn(input_sierra_path)

    # Read existing WBS WEEKLY sheet with headers at row 7 (index-based)
    wb = pd.ExcelFile(output_wbs_path)
    if "WEEKLY" not in wb.sheet_names:
        raise ValueError("WEEKLY sheet not found in WBS output")

    df = pd.read_excel(output_wbs_path, sheet_name="WEEKLY", header=7)
    df = df.dropna(how="all")
    # The first row often repeats the header literals; drop it if so
    if (df.iloc[0:1].astype(str).apply(lambda x: (x == 'Employee Name').any(), axis=1).any()):
        df = df.iloc[1:]

    # Identify canonical column names present in WBS
    name_col = "Employee Name" if "Employee Name" in df.columns else ("Unnamed: 2" if "Unnamed: 2" in df.columns else None)
    ssn_col  = "SSN"            if "SSN" in df.columns else ("Unnamed: 1" if "Unnamed: 1" in df.columns else None)
    totals_col = "Totals" if "Totals" in df.columns else None

    if name_col is None or totals_col is None:
        raise ValueError("Expected columns not found in WBS (need Employee Name and Totals)")

    # Build a lookup of current rows by name
    current_by_name = {}
    for _, row in df.iterrows():
        nm = str(row.get(name_col, "")).strip()
        if nm:
            current_by_name[nm] = row

    # Use the first real data row as a template for new zero rows
    template = df.iloc[0].copy()
    for c in df.columns:
        # zero out numerics, blank strings otherwise
        template[c] = 0 if pd.api.types.is_numeric_dtype(df[c]) else ""

    # Rebuild dataframe strictly in roster order
    new_rows = []
    for nm in roster:
        if nm in current_by_name:
            r = current_by_name[nm].copy()
        else:
            r = template.copy()
            r[name_col] = nm
            # try to backfill SSN from Sierra if we have one
            if ssn_col is not None:
                r[ssn_col] = ssn_map.get(nm, r.get(ssn_col, ""))
            # ensure totals are zero for missing
            if totals_col is not None:
                r[totals_col] = 0
        new_rows.append(r)

    new_df = pd.DataFrame(new_rows, columns=df.columns)

    # Write back, preserving other sheets
    with pd.ExcelWriter(output_wbs_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        new_df.to_excel(writer, sheet_name="WEEKLY", index=False, header=True)
