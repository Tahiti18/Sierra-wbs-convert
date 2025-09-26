#!/usr/bin/env python3
"""
Examine the actual structure of both Sierra and WBS files
"""

import pandas as pd

def examine_sierra_file():
    """Look at Sierra file structure in detail"""
    print("=== SIERRA FILE STRUCTURE ===")
    
    try:
        # Read raw data
        df = pd.read_excel("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", header=None)
        print(f"Raw shape: {df.shape}")
        print("First 15 rows:")
        for i in range(min(15, len(df))):
            print(f"Row {i}: {df.iloc[i].tolist()}")
        print()
        
        # Try different header positions
        for header_row in [0, 1, 2, 3, 4]:
            try:
                df_test = pd.read_excel("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", header=header_row)
                print(f"Header at row {header_row}: {list(df_test.columns)}")
                if 'Name' in df_test.columns or 'Employee' in str(df_test.columns):
                    print(f"Found employee data at header row {header_row}")
                    return header_row
            except:
                pass
        
    except Exception as e:
        print(f"Error: {e}")
    
    return None

def examine_wbs_file():
    """Look at WBS file structure in detail"""
    print("\n=== WBS GOLD STANDARD FILE STRUCTURE ===")
    
    try:
        # Read raw data  
        df = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=None)
        print(f"Raw shape: {df.shape}")
        print("First 20 rows:")
        for i in range(min(20, len(df))):
            row_data = df.iloc[i].tolist()
            # Look for employee names or SSNs
            row_str = str(row_data)
            if any(keyword in row_str for keyword in ['Employee', 'SSN', 'Name', 'Dianne', 'Garcia', 'Hernandez']):
                print(f"Row {i} *** EMPLOYEE DATA ***: {row_data}")
            else:
                print(f"Row {i}: {row_data}")
        print()
        
        # Try different header positions
        for header_row in range(10):
            try:
                df_test = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=header_row)
                cols = list(df_test.columns)
                if any('employee' in str(col).lower() or 'name' in str(col).lower() for col in cols):
                    print(f"Found employee columns at header row {header_row}: {cols}")
                    return header_row
            except:
                pass
                
    except Exception as e:
        print(f"Error: {e}")
    
    return None

def main():
    sierra_header = examine_sierra_file()
    wbs_header = examine_wbs_file()
    
    print(f"\n=== SUMMARY ===")
    print(f"Sierra header row: {sierra_header}")
    print(f"WBS header row: {wbs_header}")

if __name__ == "__main__":
    main()