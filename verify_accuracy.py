#!/usr/bin/env python3
"""
Verify WBS conversion accuracy with gold standard
"""

from wbs_ordered_converter import WBSOrderedConverter
import pandas as pd

def main():
    print("=== WBS CONVERSION ACCURACY VERIFICATION ===")
    
    # Initialize converter
    converter = WBSOrderedConverter()
    
    # Gold standard values from WBS_Payroll_9_12_25_for_Marwan.xlsx
    print(f"✅ WBS Master Order loaded: {len(converter.wbs_order)} employees")
    print(f"✅ First 10 WBS employees:")
    for i, emp in enumerate(converter.wbs_order[:10], 1):
        print(f"   {i:2d}. {emp}")
    
    # Test with the Sierra file from screenshot
    sierra_file = "Sierra_Payroll_9_12_25_for_Marwan_2.xlsx"
    
    try:
        print(f"\n=== PROCESSING SIERRA FILE: {sierra_file} ===")
        
        # Parse Sierra file
        employee_hours = converter.parse_sierra_file(sierra_file)
        print(f"✅ Sierra employees found: {len(employee_hours)}")
        
        # Process all WBS employees in order
        print(f"\n=== FULL WBS CONVERSION (ALL {len(converter.wbs_order)} EMPLOYEES) ===")
        
        total_amount = 0.0
        processed_count = 0
        
        for i, wbs_name in enumerate(converter.wbs_order, 1):
            emp_info = converter.find_employee_info(wbs_name)
            
            if wbs_name in employee_hours:
                hours_data = employee_hours[wbs_name]
                total_hours = hours_data['total_hours']
                rate = hours_data['rate']
                pay_calc = converter.apply_wbs_overtime_rules(total_hours, rate, wbs_name)
                amount = pay_calc['total_amount']
                total_amount += amount
                processed_count += 1
                
                if amount > 0:  # Only show employees with actual hours
                    print(f"   {i:2d}. {wbs_name:<25} SSN: {emp_info['ssn']} Amount: ${amount:,.2f}")
            else:
                # Missing employee - would show as $0.00 in WBS output
                pass
        
        print(f"\n=== SUMMARY ===")
        print(f"Total WBS employees: {len(converter.wbs_order)}")
        print(f"Sierra employees found: {processed_count}")
        print(f"Missing employees (will show as $0.00): {len(converter.wbs_order) - processed_count}")
        print(f"Total payroll amount: ${total_amount:,.2f}")
        
        # Verify key employees from gold standard
        key_employees = [
            ("Robleza, Dianne", 112.0),
            ("Garcia, Miguel A", 1160.0),
            ("Hernandez, Diego", 1840.0),
            ("Shafer, Emily", 1634.62),  # Salary employee
            ("Young, Giana L", 1538.47)  # Salary employee
        ]
        
        print(f"\n=== GOLD STANDARD VERIFICATION ===")
        for emp_name, expected in key_employees:
            if emp_name in employee_hours:
                hours_data = employee_hours[emp_name]
                pay_calc = converter.apply_wbs_overtime_rules(
                    hours_data['total_hours'], 
                    hours_data['rate'], 
                    emp_name
                )
                actual = pay_calc['total_amount']
                status = "✅ MATCH" if abs(actual - expected) < 0.01 else "❌ MISMATCH"
                print(f"   {emp_name}: Expected ${expected}, Got ${actual} {status}")
            else:
                print(f"   {emp_name}: ❌ NOT FOUND in Sierra file")
                
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()