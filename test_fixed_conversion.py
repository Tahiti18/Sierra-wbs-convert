#!/usr/bin/env python3
"""
Test the fixed conversion with proper Sierra calculations
"""

from wbs_ordered_converter import WBSOrderedConverter
import pandas as pd

def test_fixed_mappings():
    """Test that name mapping fixes work"""
    print("=== TESTING FIXED NAME MAPPINGS ===")
    
    converter = WBSOrderedConverter()
    
    # Test the problematic employees
    test_cases = [
        ("Daniel Carrasco", "Should now map to Mateos, Daniel"),
        ("Kevin Cortez", "Should now map to Cortez, Kevin"), 
        ("Kevin Duarte", "Should still map to Duarte, Kevin")
    ]
    
    for sierra_name, expected in test_cases:
        wbs_name = converter.normalize_name(sierra_name)
        print(f"  '{sierra_name}' → '{wbs_name}' ({expected})")
        
        # Check if it's in WBS order
        if wbs_name in converter.wbs_order:
            print(f"    ✅ Found in WBS order")
        else:
            print(f"    ❌ NOT in WBS order")

def test_sierra_conversion():
    """Test complete Sierra conversion with fixes"""
    print(f"\n=== TESTING COMPLETE SIERRA CONVERSION ===")
    
    converter = WBSOrderedConverter()
    
    try:
        # Parse Sierra file
        employee_hours = converter.parse_sierra_file("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx")
        print(f"✅ Parsed Sierra file: {len(employee_hours)} employees")
        
        # Create WBS output for all 79 employees
        wbs_output = []
        total_amount = 0
        mapped_count = 0
        
        for wbs_name in converter.wbs_order:
            emp_info = converter.find_employee_info(wbs_name)
            
            if wbs_name in employee_hours:
                # Employee exists in Sierra
                sierra_data = employee_hours[wbs_name]
                hours = sierra_data['total_hours']
                rate = sierra_data['rate']
                
                # Apply WBS overtime rules
                pay_calc = converter.apply_wbs_overtime_rules(hours, rate, wbs_name)
                amount = pay_calc['total_amount']
                mapped_count += 1
                
                print(f"✅ {wbs_name}: {hours}h @ ${rate}/h = ${amount:.2f}")
            else:
                # Employee missing from Sierra
                amount = 0.0
                print(f"⭕ {wbs_name}: $0.00 (missing from Sierra)")
            
            wbs_output.append({
                'employee_name': wbs_name,
                'ssn': emp_info['ssn'],
                'amount': amount
            })
            total_amount += amount
        
        print(f"\n=== CONVERSION RESULTS ===")
        print(f"Total WBS employees: {len(wbs_output)}")
        print(f"Mapped from Sierra: {mapped_count}")
        print(f"Missing from Sierra: {len(wbs_output) - mapped_count}")
        print(f"Total amount: ${total_amount:,.2f}")
        
        # Check for specific fixes
        print(f"\n=== CHECKING SPECIFIC FIXES ===")
        
        # Check Daniel Carrasco recovery
        mateos_data = next((emp for emp in wbs_output if emp['employee_name'] == 'Mateos, Daniel'), None)
        if mateos_data and mateos_data['amount'] > 0:
            print(f"✅ Daniel Carrasco recovered: Mateos, Daniel = ${mateos_data['amount']:.2f}")
        else:
            print(f"❌ Daniel Carrasco still missing")
        
        # Check Kevin separation
        cortez_data = next((emp for emp in wbs_output if emp['employee_name'] == 'Cortez, Kevin'), None)
        duarte_data = next((emp for emp in wbs_output if emp['employee_name'] == 'Duarte, Kevin'), None)
        
        if cortez_data and cortez_data['amount'] > 0:
            print(f"✅ Kevin Cortez mapped: Cortez, Kevin = ${cortez_data['amount']:.2f}")
        if duarte_data and duarte_data['amount'] > 0:
            print(f"✅ Kevin Duarte mapped: Duarte, Kevin = ${duarte_data['amount']:.2f}")
        
        # Show improvement
        return len(wbs_output), mapped_count, total_amount
        
    except Exception as e:
        print(f"❌ Conversion failed: {e}")
        import traceback
        traceback.print_exc()
        return 0, 0, 0

def main():
    print("=== TESTING FIXED SIERRA CONVERSION ===")
    print("Goal: Calculate from Sierra data only with proper mappings")
    
    test_fixed_mappings()
    total_employees, mapped_employees, total_amount = test_sierra_conversion()
    
    print(f"\n=== FINAL TEST RESULTS ===")
    if mapped_employees > 0:
        print(f"✅ SUCCESS: Fixed conversion working")
        print(f"   Total employees: {total_employees}")
        print(f"   Mapped from Sierra: {mapped_employees}")
        print(f"   Total calculated: ${total_amount:,.2f}")
        print(f"   Missing employees: {total_employees - mapped_employees} (correctly $0.00)")
        
        improvement = mapped_employees - 65  # Previous was 65 mapped
        if improvement > 0:
            print(f"🎯 IMPROVEMENT: +{improvement} employees recovered")
        
    else:
        print(f"❌ Issues remain - need further fixes")

if __name__ == "__main__":
    main()