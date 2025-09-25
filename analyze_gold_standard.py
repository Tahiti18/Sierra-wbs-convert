#!/usr/bin/env python3
"""
Quick analysis of WBS gold standard to understand exact format
"""
import pandas as pd
from openpyxl import load_workbook

def analyze_gold_standard():
    print("=== WBS GOLD STANDARD ANALYSIS ===")
    
    # Load the gold standard file
    try:
        # Load with openpyxl to see formulas
        wb = load_workbook('/home/user/webapp/wbs_gold_standard.xlsx')
        ws = wb.active
        
        print(f"Worksheet name: {ws.title}")
        print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")
        
        # Check header row (assuming row 1)
        print("\n=== HEADER ROW ===")
        header_row = []
        for col in range(1, min(30, ws.max_column + 1)):
            cell_value = ws.cell(row=1, column=col).value
            header_row.append(cell_value)
            print(f"Col {col}: {cell_value}")
            
        # Find Dianne's row (key reference employee)
        print("\n=== LOOKING FOR DIANNE ROBLEZA ===")
        dianne_row_num = None
        for row_num in range(2, min(50, ws.max_row + 1)):
            name_cell = ws.cell(row=row_num, column=3).value
            if name_cell and "Dianne" in str(name_cell):
                dianne_row_num = row_num
                print(f"Found Dianne in row {row_num}: {name_cell}")
                break
        
        if dianne_row_num:
            print(f"\n=== DIANNE'S DATA (Row {dianne_row_num}) ===")
            dianne_data = []
            for col in range(1, min(30, ws.max_column + 1)):
                cell = ws.cell(row=dianne_row_num, column=col)
                cell_value = cell.value
                
                # Check if it's a formula
                if hasattr(cell, 'data_type') and cell.data_type == 'f':
                    print(f"Col {col} ({header_row[col-1] if col <= len(header_row) else 'Unknown'}): FORMULA = {cell_value}")
                else:
                    print(f"Col {col} ({header_row[col-1] if col <= len(header_row) else 'Unknown'}): {cell_value}")
                dianne_data.append(cell_value)
        
        # Also load with pandas to see calculated values
        print("\n=== PANDAS VIEW (CALCULATED VALUES) ===")
        df = pd.read_excel('/home/user/webapp/wbs_gold_standard.xlsx', sheet_name=0)
        print(f"DataFrame shape: {df.shape}")
        print("\nColumns:")
        for i, col in enumerate(df.columns):
            print(f"  {i}: {col}")
            
        # Find Dianne in pandas
        dianne_rows = df[df.iloc[:, 2].astype(str).str.contains("Dianne", na=False)]
        if not dianne_rows.empty:
            print("\n=== DIANNE'S CALCULATED VALUES ===")
            dianne_row = dianne_rows.iloc[0]
            for i, value in enumerate(dianne_row):
                col_name = df.columns[i] if i < len(df.columns) else f"Col_{i}"
                print(f"Col {i+1} ({col_name}): {value}")
        
    except Exception as e:
        print(f"Error analyzing gold standard: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_gold_standard()