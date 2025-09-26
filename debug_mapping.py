#!/usr/bin/env python3
"""
Debug Sierra to WBS employee name mapping
"""

from wbs_ordered_converter import WBSOrderedConverter

def main():
    converter = WBSOrderedConverter()
    
    print("=== DEBUGGING EMPLOYEE MAPPING ===")
    
    # Parse Sierra file
    sierra_file = "Sierra_Payroll_9_12_25_for_Marwan_2.xlsx"
    employee_hours = converter.parse_sierra_file(sierra_file)
    
    print(f"Sierra file: {len(employee_hours)} employees")
    print(f"WBS master: {len(converter.wbs_order)} employees") 
    
    # Find Sierra employees NOT in WBS order
    sierra_not_in_wbs = []
    for sierra_name in employee_hours.keys():
        if sierra_name not in converter.wbs_order:
            sierra_not_in_wbs.append(sierra_name)
    
    print(f"\n=== SIERRA EMPLOYEES NOT FOUND IN WBS ORDER ({len(sierra_not_in_wbs)}) ===")
    for name in sierra_not_in_wbs:
        hours_data = employee_hours[name]
        print(f"  {name} | ${hours_data['total_hours'] * hours_data['rate']:.2f}")
    
    # Count WBS employees with payroll data
    wbs_with_data = []
    wbs_without_data = []
    
    for wbs_name in converter.wbs_order:
        if wbs_name in employee_hours:
            wbs_with_data.append(wbs_name)
        else:
            wbs_without_data.append(wbs_name)
    
    print(f"\n=== WBS EMPLOYEES WITH PAYROLL DATA ({len(wbs_with_data)}) ===")
    total_payroll = 0
    for name in wbs_with_data[:10]:  # First 10
        hours_data = employee_hours[name]
        pay_calc = converter.apply_wbs_overtime_rules(
            hours_data['total_hours'], 
            hours_data['rate'], 
            name
        )
        amount = pay_calc['total_amount']
        total_payroll += amount
        print(f"  {name} | ${amount:.2f}")
    
    if len(wbs_with_data) > 10:
        print(f"  ... and {len(wbs_with_data) - 10} more")
    
    # Calculate total for ALL WBS employees with data
    full_total = 0
    for name in wbs_with_data:
        hours_data = employee_hours[name]
        pay_calc = converter.apply_wbs_overtime_rules(
            hours_data['total_hours'], 
            hours_data['rate'], 
            name
        )
        full_total += pay_calc['total_amount']
    
    print(f"\n=== SUMMARY ===")
    print(f"WBS employees with payroll data: {len(wbs_with_data)}")
    print(f"WBS employees with $0.00: {len(wbs_without_data)}")
    print(f"Sierra employees not mapping to WBS: {len(sierra_not_in_wbs)}")
    print(f"Total payroll amount: ${full_total:,.2f}")
    
    print(f"\n=== WBS EMPLOYEES WITH $0.00 (First 10) ===")
    for name in wbs_without_data[:10]:
        print(f"  {name}")
    if len(wbs_without_data) > 10:
        print(f"  ... and {len(wbs_without_data) - 10} more")

if __name__ == "__main__":
    main()