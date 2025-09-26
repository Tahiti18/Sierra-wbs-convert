#!/usr/bin/env python3

import pandas as pd
import numpy as np
from datetime import datetime

def analyze_file_structure(file_path, file_name):
    """Analyze Excel file structure"""
    print(f"\n{'='*80}")
    print(f"ANALYZING: {file_name}")
    print(f"{'='*80}")
    
    try:
        # Try to read the Excel file
        xl_file = pd.ExcelFile(file_path)
        print(f"Sheet names: {xl_file.sheet_names}")
        
        # Analyze each sheet
        for sheet_name in xl_file.sheet_names:
            print(f"\n--- SHEET: {sheet_name} ---")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"Shape: {df.shape}")
            print(f"Columns: {list(df.columns)}")
            
            # Show first few rows
            print("\nFirst 5 rows:")
            print(df.head())
            
            # Look for employee names and key data
            if not df.empty:
                # Find columns that might contain employee names
                name_columns = []
                for col in df.columns:
                    col_str = str(col).lower()
                    if any(word in col_str for word in ['name', 'employee', 'emp']):
                        name_columns.append(col)
                
                if name_columns:
                    print(f"\nPotential name columns: {name_columns}")
                    for col in name_columns:
                        unique_vals = df[col].dropna().unique()
                        print(f"{col} - Unique values (first 10): {unique_vals[:10]}")
                
                # Find numeric columns (hours, rates, amounts)
                numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
                if numeric_columns:
                    print(f"\nNumeric columns: {numeric_columns}")
                    for col in numeric_columns:
                        print(f"{col}: min={df[col].min():.2f}, max={df[col].max():.2f}, sum={df[col].sum():.2f}")
            
    except Exception as e:
        print(f"Error reading {file_name}: {e}")

def main():
    print("Sierra Payroll to WBS Format Analysis")
    print("=" * 80)
    
    # Analyze the new files
    analyze_file_structure("/home/user/webapp/WBS_Payroll_9_12_25_for_Marwan_2.xlsx", "WBS Payroll (TARGET FORMAT)")
    analyze_file_structure("/home/user/webapp/Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", "Sierra Payroll (SOURCE FORMAT)")
    
    print(f"\n{'='*80}")
    print("SUMMARY:")
    print("- WBS file is the TARGET format we need to match exactly")
    print("- Sierra file is the SOURCE format we need to convert FROM")
    print("- All calculations, employee order, formatting must match WBS exactly")
    print(f"{'='*80}")

if __name__ == "__main__":
    main()