#!/usr/bin/env python3
"""
Exact comparison between Sierra and WBS files using proper parsing
"""

import pandas as pd
from wbs_ordered_converter import WBSOrderedConverter
import re

def parse_sierra_file():
    """Parse Sierra file correctly"""
    print("=== PARSING SIERRA FILE ===")
    
    df = pd.read_excel("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", header=0)
    
    # Group by employee name and sum hours
    employee_totals = {}
    
    for _, row in df.iterrows():
        name = row.get('Name')
        hours = row.get('Hours')
        rate = row.get('Rate')
        
        if pd.notna(name) and pd.notna(hours) and pd.notna(rate):
            name = str(name).strip()
            hours = float(hours)
            rate = float(rate)
            
            # Skip header rows and invalid data
            if name == 'Name' or hours == 0 or rate == 0:
                continue
            
            if name not in employee_totals:
                employee_totals[name] = {'hours': 0, 'rate': rate}
            
            employee_totals[name]['hours'] += hours
            # Use the rate from the row (should be consistent per employee)
            employee_totals[name]['rate'] = rate
    
    # Calculate total amounts
    for name in employee_totals:
        emp_data = employee_totals[name]
        emp_data['amount'] = emp_data['hours'] * emp_data['rate']
    
    print(f"Found {len(employee_totals)} employees in Sierra file")
    return employee_totals

def parse_wbs_file():
    """Parse WBS gold standard file correctly"""
    print("\n=== PARSING WBS GOLD STANDARD FILE ===")
    
    # Read raw data and find the correct header row
    df_raw = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=None)
    
    # Find header row (contains 'SSN', 'Employee Name', etc.)
    header_row = None
    for i in range(len(df_raw)):
        row_data = df_raw.iloc[i].tolist()
        if any('SSN' in str(cell) for cell in row_data) and any('Employee Name' in str(cell) for cell in row_data):
            header_row = i
            break
    
    if header_row is None:
        print("Could not find header row in WBS file")
        return {}
    
    print(f"Found WBS header at row {header_row}")
    
    # Read with correct header
    df = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=header_row)
    
    employees = {}
    
    for _, row in df.iterrows():
        ssn = row.get('SSN')
        name = row.get('Employee Name')
        total_amount = row.get('Totals')  # Last column should be totals
        
        if pd.notna(ssn) and pd.notna(name) and pd.notna(total_amount):
            ssn = str(ssn).strip()
            name = str(name).strip()
            
            # Skip non-employee rows
            if ssn == 'SSN' or name == 'Employee Name' or name == 'Totals':
                continue
            
            try:
                amount = float(total_amount)
                employees[name] = {
                    'ssn': ssn,
                    'amount': amount,
                    'rate': row.get('Pay Rate', 0),
                    'regular_hours': row.get('A01', 0),  # Regular hours column
                    'overtime_hours': row.get('A02', 0)  # Overtime hours column
                }
            except (ValueError, TypeError):
                continue
    
    print(f"Found {len(employees)} employees in WBS gold standard")
    return employees

def main():
    print("=== EXACT SIERRA VS WBS COMPARISON ===")
    
    # Parse both files
    sierra_employees = parse_sierra_file()
    wbs_employees = parse_wbs_file()
    
    if not sierra_employees or not wbs_employees:
        print("Failed to parse files")
        return
    
    # Initialize converter for name normalization
    converter = WBSOrderedConverter()
    
    # Normalize Sierra names to match WBS format
    sierra_normalized = {}
    for sierra_name, sierra_data in sierra_employees.items():
        # Convert "First Last" to "Last, First" format
        normalized_name = converter.normalize_name(sierra_name)
        sierra_normalized[normalized_name] = sierra_data
    
    print(f"\n=== COMPARISON RESULTS ===")
    print(f"Sierra employees (normalized): {len(sierra_normalized)}")
    print(f"WBS employees: {len(wbs_employees)}")
    
    # Compare each WBS employee
    perfect_matches = []
    amount_mismatches = []
    missing_in_sierra = []
    
    total_wbs_amount = 0
    total_sierra_amount = 0
    
    print(f"\n=== DETAILED EMPLOYEE COMPARISON ===")
    
    for wbs_name, wbs_data in wbs_employees.items():
        wbs_amount = wbs_data['amount']
        total_wbs_amount += wbs_amount
        
        if wbs_name in sierra_normalized:
            sierra_data = sierra_normalized[wbs_name]
            sierra_amount = sierra_data['amount']
            total_sierra_amount += sierra_amount
            
            difference = sierra_amount - wbs_amount
            
            if abs(difference) < 0.01:  # Match within 1 cent
                perfect_matches.append(wbs_name)
                print(f"✅ {wbs_name}: ${wbs_amount:.2f}")
            else:
                amount_mismatches.append({
                    'name': wbs_name,
                    'wbs_amount': wbs_amount,
                    'sierra_amount': sierra_amount,
                    'difference': difference
                })
                print(f"❌ {wbs_name}: WBS=${wbs_amount:.2f}, Sierra=${sierra_amount:.2f}, Diff=${difference:+.2f}")
        else:
            missing_in_sierra.append({
                'name': wbs_name,
                'amount': wbs_amount
            })
            print(f"⚠️  {wbs_name}: ${wbs_amount:.2f} (MISSING FROM SIERRA)")
    
    # Check for extra Sierra employees
    extra_in_sierra = []
    for sierra_name in sierra_normalized:
        if sierra_name not in wbs_employees:
            extra_in_sierra.append(sierra_name)
    
    print(f"\n=== FINAL RESULTS ===")
    print(f"✅ Perfect matches: {len(perfect_matches)}")
    print(f"❌ Amount mismatches: {len(amount_mismatches)}")
    print(f"⚠️  Missing in Sierra: {len(missing_in_sierra)}")
    print(f"⚠️  Extra in Sierra: {len(extra_in_sierra)}")
    print(f"WBS Total: ${total_wbs_amount:,.2f}")
    print(f"Sierra Total: ${total_sierra_amount:,.2f}")
    print(f"Difference: ${total_sierra_amount - total_wbs_amount:+,.2f}")
    
    # Show first few mismatches for analysis
    if amount_mismatches:
        print(f"\n=== TOP AMOUNT MISMATCHES ===")
        for mismatch in sorted(amount_mismatches, key=lambda x: abs(x['difference']), reverse=True)[:10]:
            print(f"  {mismatch['name']}: ${mismatch['difference']:+,.2f}")
    
    # Show missing employees
    if missing_in_sierra:
        print(f"\n=== MISSING IN SIERRA (HIGH VALUE) ===")
        missing_sorted = sorted(missing_in_sierra, key=lambda x: x['amount'], reverse=True)
        for missing in missing_sorted[:10]:
            print(f"  {missing['name']}: ${missing['amount']:.2f}")

if __name__ == "__main__":
    main()