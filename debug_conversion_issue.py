#!/usr/bin/env python3

import pandas as pd
from wbs_ordered_converter import WBSOrderedConverter

def debug_conversion():
    """Debug the conversion issue step by step"""
    
    print("DEBUGGING CONVERSION ISSUE")
    print("=" * 80)
    
    # Initialize converter
    converter = WBSOrderedConverter()
    
    # Parse Sierra file
    print("Step 1: Parsing Sierra file...")
    employee_hours = converter.parse_sierra_file("/home/user/webapp/Sierra_Payroll_9_12_25_for_Marwan_2.xlsx")
    
    print(f"Found {len(employee_hours)} employees with hours")
    print("\nFirst few employees:")
    for i, (name, data) in enumerate(list(employee_hours.items())[:5]):
        print(f"  {name}: {data['total_hours']} hours @ ${data['rate']}/hr = ${data['total_hours'] * data['rate']}")
    
    # Test overtime calculations
    print("\nStep 2: Testing overtime calculations...")
    test_employee = list(employee_hours.items())[0]
    name, hours_data = test_employee
    
    print(f"Testing {name}: {hours_data['total_hours']} hours @ ${hours_data['rate']}/hr")
    
    # Apply WBS overtime rules
    pay_calc = converter.apply_wbs_overtime_rules(hours_data['total_hours'], hours_data['rate'], name)
    print(f"Overtime calculation result: {pay_calc}")
    
    # Check employee info mapping
    print("\nStep 3: Testing employee info mapping...")
    emp_info = converter.find_employee_info(name)
    print(f"Employee info for {name}: {emp_info}")
    
    # Compare with actual WBS to see what Dianne should have
    print("\nStep 4: Checking Dianne specifically...")
    actual_wbs = pd.read_excel("/home/user/webapp/WBS_Payroll_9_12_25_for_Marwan_2.xlsx", sheet_name="WEEKLY")
    
    # Find Dianne in actual WBS
    for idx, row in actual_wbs.iterrows():
        if 'Dianne' in str(row.iloc[2]):
            print(f"Dianne in actual WBS:")
            print(f"  Row {idx}: Name={row.iloc[2]}, Hours={row.iloc[7]}, Total=${row.iloc[27]}")
            break
    
    # Check if Dianne is in our Sierra data
    if "Dianne Robleza" in [name for name in employee_hours.keys()]:
        dianne_data = employee_hours["Dianne Robleza"]
        print(f"Dianne in our Sierra data: {dianne_data}")
    else:
        # Try to find Dianne with different name format
        dianne_matches = [name for name in employee_hours.keys() if 'Dianne' in name]
        print(f"Dianne name matches in Sierra: {dianne_matches}")

if __name__ == "__main__":
    debug_conversion()