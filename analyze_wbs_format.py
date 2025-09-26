#!/usr/bin/env python3

import pandas as pd
import numpy as np

def analyze_wbs_detailed():
    """Detailed analysis of WBS payroll format"""
    
    # Read WBS file
    df = pd.read_excel("/home/user/webapp/WBS_Payroll_9_12_25_for_Marwan_2.xlsx", sheet_name="WEEKLY")
    
    print("WBS PAYROLL FORMAT ANALYSIS")
    print("=" * 80)
    
    # Find actual data rows (skip headers)
    print("\nFinding actual employee data rows...")
    
    # Look at different row ranges
    for i in range(min(20, len(df))):
        row_data = df.iloc[i]
        name_col = row_data.iloc[1]  # Second column usually has employee info
        print(f"Row {i}: {name_col}")
    
    print("\n" + "=" * 80)
    
    # Try to find employee rows by looking for SSNs or employee names
    employee_rows = []
    for idx, row in df.iterrows():
        # Check if this row contains employee data
        # Look at column 1 (DO NOT EDIT column) for SSNs or employee indicators
        col1_val = str(row.iloc[1]).strip()
        
        # Skip header rows
        if any(skip in col1_val.upper() for skip in ['CLIENT', 'PERIOD', 'CHECK', 'REPORT', 'VERSION', 'DEPT']):
            continue
            
        # Look for SSN pattern or employee indicators
        if len(col1_val) >= 3 and col1_val not in ['nan', 'NaN', '']:
            # Check if it looks like an employee row
            if any(c.isdigit() for c in col1_val) or col1_val.isalpha():
                employee_rows.append((idx, col1_val, row))
    
    print(f"Found {len(employee_rows)} potential employee rows:")
    
    for idx, identifier, row in employee_rows[:10]:  # Show first 10
        print(f"Row {idx}: {identifier}")
        # Show all non-null values in this row
        non_null = [(i, v) for i, v in enumerate(row) if pd.notna(v) and str(v).strip() != '']
        print(f"  Non-null values: {non_null}")
        print()

def analyze_sierra_detailed():
    """Detailed analysis of Sierra payroll format"""
    
    # Read Sierra file
    df = pd.read_excel("/home/user/webapp/Sierra_Payroll_9_12_25_for_Marwan_2.xlsx")
    
    print("\n" + "=" * 80)
    print("SIERRA PAYROLL FORMAT ANALYSIS") 
    print("=" * 80)
    
    # Look at data structure
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    
    # Find actual employee data (rows with names and hours)
    employee_data = df.dropna(subset=['Name', 'Hours'])
    
    # Filter out header/summary rows
    employee_data = employee_data[
        (employee_data['Name'].str.len() > 3) &
        (~employee_data['Name'].str.contains('Week|Gross|Total', case=False, na=False))
    ]
    
    print(f"\nFound {len(employee_data)} employee records")
    
    # Group by employee and sum hours
    employee_totals = employee_data.groupby('Name').agg({
        'Hours': 'sum',
        'Rate': 'first',  # Assuming consistent rate per employee
        'Total': 'sum'
    }).reset_index()
    
    print("\nEmployee totals:")
    print(employee_totals.head(10))
    
    print(f"\nTotal hours across all employees: {employee_totals['Hours'].sum()}")
    print(f"Total amount: ${employee_totals['Total'].sum():.2f}")
    
    # Show unique employees
    unique_employees = sorted(employee_totals['Name'].tolist())
    print(f"\nUnique employees ({len(unique_employees)}):")
    for i, name in enumerate(unique_employees):
        print(f"{i+1:2d}. {name}")

if __name__ == "__main__":
    analyze_wbs_detailed()
    analyze_sierra_detailed()