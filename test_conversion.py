#!/usr/bin/env python3
"""
Test conversion accuracy against gold standard
"""

from wbs_ordered_converter import WBSOrderedConverter

def main():
    # Initialize converter
    converter = WBSOrderedConverter()
    
    # Gold standard totals for verification
    gold_standard_totals = {
        "Robleza, Dianne": 112.0,
        "Shafer, Emily": 1634.62,
        "Young, Giana L": 1538.47,
        "Garcia, Miguel A": 1160.0,
        "Garcia, Bryan": 0.0,
        "Hernandez, Diego": 1840.0
    }
    
    print("=== Testing WBS Conversion Accuracy ===")
    print(f"WBS Master Order: {len(converter.wbs_order)} employees")
    
    # Test Sierra file parsing
    try:
        employee_hours = converter.parse_sierra_file('Sierra_Payroll_9_12_25_for_Marwan_2.xlsx')
        print(f"✅ Sierra file parsed: {len(employee_hours)} employees found")
        
        # Test specific employees
        print("\n=== Checking Gold Standard Matches ===")
        for emp_name, expected_total in gold_standard_totals.items():
            if emp_name in employee_hours:
                hours_data = employee_hours[emp_name]
                total_hours = hours_data['total_hours']
                rate = hours_data['rate']
                
                # Calculate WBS totals
                pay_calc = converter.apply_wbs_overtime_rules(total_hours, rate, emp_name)
                calculated_total = pay_calc['total_amount']
                
                match_status = "✅ MATCH" if abs(calculated_total - expected_total) < 0.01 else "❌ MISMATCH"
                
                print(f"{emp_name}:")
                print(f"  Expected: ${expected_total}")
                print(f"  Calculated: ${calculated_total}")
                print(f"  Status: {match_status}")
                print()
            else:
                print(f"❌ {emp_name}: NOT FOUND in Sierra file")
        
        print("=== WBS Order Check ===")
        print("First 10 employees in WBS order:")
        for i, emp_name in enumerate(converter.wbs_order[:10], 1):
            print(f"  {i:2d}. {emp_name}")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()