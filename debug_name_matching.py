#!/usr/bin/env python3

from wbs_ordered_converter import WBSOrderedConverter

def debug_name_matching():
    """Debug why names aren't matching properly in the converter"""
    
    print("DEBUGGING NAME MATCHING IN CONVERTER")
    print("=" * 80)
    
    # Initialize converter
    converter = WBSOrderedConverter()
    
    # Get Sierra employee hours
    employee_hours = converter.parse_sierra_file("/home/user/webapp/Sierra_Payroll_9_12_25_for_Marwan_2.xlsx")
    
    print(f"Sierra employees parsed: {len(employee_hours)}")
    
    # Check first few employees from Sierra
    print("\nFirst 10 Sierra employees with hours:")
    for i, (name, data) in enumerate(list(employee_hours.items())[:10]):
        print(f"{i+1:2d}. '{name}' -> {data['total_hours']} hours @ ${data['rate']}/hr")
        
        # Check if this name is in WBS order
        if name in converter.wbs_order:
            wbs_index = converter.wbs_order.index(name)
            print(f"     ✅ Found in WBS at position {wbs_index}")
        else:
            print(f"     ❌ NOT found in WBS master order")
            
            # Try to find similar names
            similar = [wbs_name for wbs_name in converter.wbs_order if name.lower() in wbs_name.lower() or wbs_name.lower() in name.lower()]
            if similar:
                print(f"     Similar WBS names: {similar}")
    
    # Check why some employees show up as None in our conversion
    print(f"\n" + "=" * 60)
    print("CHECKING WBS ORDER PROCESSING")
    print("=" * 60)
    
    # Check first 10 WBS employees 
    print("First 10 WBS employees and their status:")
    for i, wbs_name in enumerate(converter.wbs_order[:10]):
        print(f"{i+1:2d}. '{wbs_name}'")
        
        if wbs_name in employee_hours:
            data = employee_hours[wbs_name]
            print(f"     ✅ Has hours: {data['total_hours']} @ ${data['rate']}/hr")
        else:
            print(f"     ❌ No hours data (employee didn't work this week)")

def test_specific_employees():
    """Test specific employees that should have data"""
    
    print(f"\n" + "=" * 60)  
    print("TESTING SPECIFIC EMPLOYEES")
    print("=" * 60)
    
    converter = WBSOrderedConverter()
    employee_hours = converter.parse_sierra_file("/home/user/webapp/Sierra_Payroll_9_12_25_for_Marwan_2.xlsx")
    
    # Test employees we know should be there
    test_names = ["Robleza, Dianne", "Stokes, Symone", "Hernandez, Diego", "Garcia, Miguel", "Shafer, Emily"]
    
    for name in test_names:
        print(f"\nTesting '{name}':")
        
        if name in employee_hours:
            data = employee_hours[name]
            print(f"  ✅ Sierra data: {data['total_hours']} hours @ ${data['rate']}/hr = ${data['total_hours'] * data['rate']}")
        else:
            print(f"  ❌ No Sierra data found")
        
        if name in converter.wbs_order:
            wbs_pos = converter.wbs_order.index(name)
            print(f"  ✅ WBS position: {wbs_pos}")
        else:
            print(f"  ❌ Not in WBS order")
        
        # Check employee info lookup
        emp_info = converter.find_employee_info(name)
        print(f"  Employee info: {emp_info.get('employee_number', 'Not found')}")

if __name__ == "__main__":
    debug_name_matching()
    test_specific_employees()