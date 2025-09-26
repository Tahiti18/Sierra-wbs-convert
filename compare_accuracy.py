#!/usr/bin/env python3

import pandas as pd
import numpy as np

def compare_wbs_formats():
    """Compare our generated WBS with the actual WBS format"""
    
    print("COMPARING WBS FORMATS")
    print("=" * 80)
    
    # Read actual WBS (target)
    actual_wbs = pd.read_excel("/home/user/webapp/WBS_Payroll_9_12_25_for_Marwan_2.xlsx", sheet_name="WEEKLY")
    
    # Read our generated WBS
    our_wbs = pd.read_excel("/home/user/webapp/test_new_conversion.xlsx", sheet_name="WEEKLY")
    
    print(f"Actual WBS shape: {actual_wbs.shape}")
    print(f"Our WBS shape: {our_wbs.shape}")
    
    # Find employee rows in both files
    def find_employee_rows(df):
        employees = []
        for idx, row in df.iterrows():
            # Look for SSN patterns (9 digits)
            col1_val = str(row.iloc[1]).strip()
            if col1_val.isdigit() and len(col1_val) == 9:
                employee_name = str(row.iloc[2]) if len(row) > 2 else "Unknown"
                employees.append({
                    'row': idx,
                    'ssn': col1_val,
                    'name': employee_name,
                    'hours': row.iloc[7] if len(row) > 7 and pd.notna(row.iloc[7]) else 0,
                    'total': row.iloc[27] if len(row) > 27 and pd.notna(row.iloc[27]) else 0
                })
        return employees
    
    actual_employees = find_employee_rows(actual_wbs)
    our_employees = find_employee_rows(our_wbs)
    
    print(f"\nActual WBS employees found: {len(actual_employees)}")
    print(f"Our WBS employees found: {len(our_employees)}")
    
    # Compare specific employees
    print("\nACTUAL WBS EMPLOYEES (first 10):")
    for i, emp in enumerate(actual_employees[:10]):
        print(f"{i+1:2d}. {emp['name']} - Hours: {emp['hours']} - Total: ${emp['total']}")
    
    print("\nOUR WBS EMPLOYEES (first 10):")
    for i, emp in enumerate(our_employees[:10]):
        print(f"{i+1:2d}. {emp['name']} - Hours: {emp['hours']} - Total: ${emp['total']}")
    
    # Find matching employees
    print("\nMATCHING ANALYSIS:")
    actual_names = {emp['name'] for emp in actual_employees}
    our_names = {emp['name'] for emp in our_employees}
    
    print(f"Names in actual but not in ours: {actual_names - our_names}")
    print(f"Names in ours but not in actual: {our_names - actual_names}")
    
    # Compare totals
    actual_total_hours = sum(emp['hours'] for emp in actual_employees if isinstance(emp['hours'], (int, float)))
    our_total_hours = sum(emp['hours'] for emp in our_employees if isinstance(emp['hours'], (int, float)))
    
    actual_total_amount = sum(emp['total'] for emp in actual_employees if isinstance(emp['total'], (int, float)))
    our_total_amount = sum(emp['total'] for emp in our_employees if isinstance(emp['total'], (int, float)))
    
    print(f"\nTOTALS COMPARISON:")
    print(f"Actual total hours: {actual_total_hours}")
    print(f"Our total hours: {our_total_hours}")
    print(f"Hours difference: {our_total_hours - actual_total_hours}")
    
    print(f"Actual total amount: ${actual_total_amount}")
    print(f"Our total amount: ${our_total_amount}")
    print(f"Amount difference: ${our_total_amount - actual_total_amount}")

if __name__ == "__main__":
    compare_wbs_formats()