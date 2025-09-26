#!/usr/bin/env python3
"""
Verify Sierra to WBS conversion accuracy using actual data
"""

from wbs_ordered_converter import WBSOrderedConverter
import pandas as pd

def main():
    print("=== VERIFYING SIERRA TO WBS CONVERSION ACCURACY ===")
    
    converter = WBSOrderedConverter()
    
    print(f"✅ WBS Master Order: {len(converter.wbs_order)} employees")
    print("First 10 WBS employees:")
    for i, name in enumerate(converter.wbs_order[:10], 1):
        print(f"  {i:2d}. {name}")
    
    # Test with the actual file from the screenshot
    sierra_file = "Sierra_Payroll_9_12_25_for_Marwan_2.xlsx"
    
    try:
        print(f"\n=== PROCESSING {sierra_file} ===")
        
        # Parse Sierra file 
        employee_hours = converter.parse_sierra_file(sierra_file)
        print(f"✅ Sierra employees found: {len(employee_hours)}")
        
        # Create complete WBS data (ALL 79 employees)
        wbs_complete = []
        total_amount = 0.0
        
        for position, wbs_name in enumerate(converter.wbs_order, 1):
            emp_info = converter.find_employee_info(wbs_name)
            
            if wbs_name in employee_hours:
                hours_data = employee_hours[wbs_name]
                total_hours = hours_data['total_hours']
                rate = hours_data['rate']
                pay_calc = converter.apply_wbs_overtime_rules(total_hours, rate, wbs_name)
                amount = pay_calc['total_amount']
                total_amount += amount
                
                wbs_complete.append({
                    'position': position,
                    'name': wbs_name,
                    'ssn': emp_info['ssn'],
                    'hours': total_hours,
                    'rate': rate,
                    'amount': amount,
                    'found_in_sierra': True
                })
            else:
                # Missing employee - zero amount
                wbs_complete.append({
                    'position': position,
                    'name': wbs_name, 
                    'ssn': emp_info['ssn'],
                    'hours': 0.0,
                    'rate': 0.0,
                    'amount': 0.0,
                    'found_in_sierra': False
                })
        
        print(f"\n=== WBS COMPLETE OUTPUT VERIFICATION ===")
        print(f"Total WBS employees: {len(wbs_complete)}")
        print(f"Employees with data: {len([e for e in wbs_complete if e['found_in_sierra']])}")
        print(f"Employees with zero: {len([e for e in wbs_complete if not e['found_in_sierra']])}")
        print(f"Total amount: ${total_amount:,.2f}")
        
        print(f"\n=== FIRST 10 WBS EMPLOYEES (SSN, NAME, AMOUNT) ===")
        for emp in wbs_complete[:10]:
            status = "✅" if emp['found_in_sierra'] else "⭕"
            print(f"{status} {emp['position']:2d}. {emp['ssn']} | {emp['name']} | ${emp['amount']:,.2f}")
            
        print(f"\n=== EMPLOYEES WITH HIGHEST AMOUNTS ===")
        high_earners = sorted([e for e in wbs_complete if e['amount'] > 0], 
                             key=lambda x: x['amount'], reverse=True)[:5]
        for emp in high_earners:
            print(f"  {emp['position']:2d}. {emp['ssn']} | {emp['name']} | ${emp['amount']:,.2f}")
        
        # Check specific employees mentioned in gold standard
        print(f"\n=== GOLD STANDARD VERIFICATION ===")
        test_cases = [
            ("Robleza, Dianne", 112.0),
            ("Garcia, Miguel A", 1160.0), 
            ("Hernandez, Diego", 1840.0),
            ("Shafer, Emily", 1634.62),
            ("Young, Giana L", 1538.47)
        ]
        
        for test_name, expected in test_cases:
            found_emp = next((e for e in wbs_complete if e['name'] == test_name), None)
            if found_emp:
                if found_emp['found_in_sierra']:
                    match = abs(found_emp['amount'] - expected) < 0.01
                    status = "✅ MATCH" if match else "❌ MISMATCH"
                    print(f"{status} {test_name}: Expected ${expected}, Got ${found_emp['amount']:.2f}")
                else:
                    print(f"⭕ {test_name}: Not in Sierra file (${expected} expected)")
            else:
                print(f"❌ {test_name}: Not in WBS order!")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()