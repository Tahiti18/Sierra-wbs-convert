#!/usr/bin/env python3
"""
Detailed analysis of Sierra vs WBS gold standard files
Find exact discrepancies and fix them systematically
"""

import pandas as pd
from wbs_ordered_converter import WBSOrderedConverter

def analyze_sierra_file(filename):
    """Analyze Sierra file structure and extract all employee data"""
    print(f"=== ANALYZING SIERRA FILE: {filename} ===")
    
    try:
        df = pd.read_excel(filename, header=0)
        print(f"Columns: {list(df.columns)}")
        print(f"Rows: {len(df)}")
        print(f"Sample data:")
        print(df.head())
        
        # Extract employee data
        employees = {}
        for _, row in df.iterrows():
            if pd.notna(row.get('Name', '')):
                name = str(row['Name']).strip()
                if name and name != 'Name':
                    total_hours = float(row.get('Total Hours', 0) or 0)
                    rate = float(row.get('Rate', 0) or 0)
                    amount = total_hours * rate
                    
                    employees[name] = {
                        'hours': total_hours,
                        'rate': rate,
                        'amount': amount
                    }
        
        print(f"Extracted {len(employees)} employees from Sierra file")
        return employees
        
    except Exception as e:
        print(f"Error analyzing Sierra file: {e}")
        return {}

def analyze_wbs_file(filename):
    """Analyze WBS gold standard file"""
    print(f"\n=== ANALYZING WBS GOLD STANDARD: {filename} ===")
    
    try:
        df = pd.read_excel(filename, header=0)
        print(f"Columns: {list(df.columns)}")
        print(f"Rows: {len(df)}")
        print(f"Sample data:")
        print(df.head())
        
        # Extract employee data
        employees = {}
        for _, row in df.iterrows():
            name = str(row.get('Employee Name', '')).strip()
            if name and name != 'Employee Name' and name != 'Totals':
                # Get the total amount
                total_col = None
                for col in df.columns:
                    if 'total' in col.lower() or 'amount' in col.lower():
                        total_col = col
                        break
                
                if total_col:
                    amount = float(row.get(total_col, 0) or 0)
                    employees[name] = {
                        'amount': amount,
                        'ssn': str(row.get('SSN', '')).strip(),
                        'hours': float(row.get('Hours', 0) or 0),
                        'rate': float(row.get('Rate', 0) or 0)
                    }
        
        print(f"Extracted {len(employees)} employees from WBS gold standard")
        return employees
        
    except Exception as e:
        print(f"Error analyzing WBS file: {e}")
        return {}

def main():
    print("=== DETAILED SIERRA vs WBS ANALYSIS ===")
    
    # Analyze both files
    sierra_data = analyze_sierra_file("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx")
    wbs_data = analyze_wbs_file("WBS_Payroll_9_12_25_for_Marwan.xlsx")
    
    if not sierra_data or not wbs_data:
        print("Could not load data from files")
        return
    
    # Initialize converter for name mapping
    converter = WBSOrderedConverter()
    
    print(f"\n=== DATA COMPARISON ===")
    print(f"Sierra employees: {len(sierra_data)}")
    print(f"WBS gold employees: {len(wbs_data)}")
    
    # Map Sierra names to WBS names
    sierra_mapped = {}
    for sierra_name, sierra_info in sierra_data.items():
        wbs_name = converter.normalize_name(sierra_name)
        sierra_mapped[wbs_name] = sierra_info
    
    print(f"Sierra employees after mapping: {len(sierra_mapped)}")
    
    # Find matches and mismatches
    perfect_matches = []
    amount_mismatches = []
    missing_in_sierra = []
    extra_in_sierra = []
    
    # Check each WBS employee
    for wbs_name, wbs_info in wbs_data.items():
        wbs_amount = wbs_info['amount']
        
        if wbs_name in sierra_mapped:
            sierra_info = sierra_mapped[wbs_name]
            sierra_amount = sierra_info['amount']
            
            if abs(sierra_amount - wbs_amount) < 0.01:
                perfect_matches.append({
                    'name': wbs_name,
                    'amount': wbs_amount
                })
            else:
                amount_mismatches.append({
                    'name': wbs_name,
                    'wbs_amount': wbs_amount,
                    'sierra_amount': sierra_amount,
                    'difference': sierra_amount - wbs_amount,
                    'sierra_hours': sierra_info['hours'],
                    'sierra_rate': sierra_info['rate']
                })
        else:
            missing_in_sierra.append({
                'name': wbs_name,
                'wbs_amount': wbs_amount
            })
    
    # Check for extra Sierra employees
    for sierra_name in sierra_mapped:
        if sierra_name not in wbs_data:
            extra_in_sierra.append({
                'name': sierra_name,
                'sierra_amount': sierra_mapped[sierra_name]['amount']
            })
    
    # Print detailed results
    print(f"\n=== DETAILED RESULTS ===")
    print(f"✅ Perfect matches: {len(perfect_matches)}")
    print(f"❌ Amount mismatches: {len(amount_mismatches)}")
    print(f"⚠️  Missing in Sierra: {len(missing_in_sierra)}")
    print(f"⚠️  Extra in Sierra: {len(extra_in_sierra)}")
    
    if amount_mismatches:
        print(f"\n=== ❌ AMOUNT MISMATCHES ({len(amount_mismatches)}) ===")
        for mismatch in amount_mismatches[:10]:
            print(f"  {mismatch['name']}")
            print(f"    WBS: ${mismatch['wbs_amount']:,.2f}")
            print(f"    Sierra: ${mismatch['sierra_amount']:,.2f} ({mismatch['sierra_hours']}h @ ${mismatch['sierra_rate']}/h)")
            print(f"    Diff: ${mismatch['difference']:+,.2f}")
            print()
    
    if missing_in_sierra:
        print(f"\n=== ⚠️ MISSING IN SIERRA ({len(missing_in_sierra)}) ===")
        total_missing = 0
        for missing in missing_in_sierra[:10]:
            print(f"  {missing['name']} → ${missing['wbs_amount']:,.2f}")
            total_missing += missing['wbs_amount']
        if len(missing_in_sierra) > 10:
            for missing in missing_in_sierra[10:]:
                total_missing += missing['wbs_amount']
            print(f"  ... and {len(missing_in_sierra) - 10} more")
        print(f"  Total missing: ${total_missing:,.2f}")
    
    if extra_in_sierra:
        print(f"\n=== ⚠️ EXTRA IN SIERRA ({len(extra_in_sierra)}) ===")
        for extra in extra_in_sierra:
            print(f"  {extra['name']} → ${extra['sierra_amount']:,.2f}")
    
    # Calculate totals
    wbs_total = sum(emp['amount'] for emp in wbs_data.values())
    sierra_total = sum(emp['amount'] for emp in sierra_data.values())
    
    print(f"\n=== TOTALS ===")
    print(f"WBS Gold Standard: ${wbs_total:,.2f}")
    print(f"Sierra File: ${sierra_total:,.2f}")
    print(f"Difference: ${sierra_total - wbs_total:+,.2f}")

if __name__ == "__main__":
    main()