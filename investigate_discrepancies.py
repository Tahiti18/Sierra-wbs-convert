#!/usr/bin/env python3
"""
Investigate why calculated amounts differ between Sierra and WBS
"""

import pandas as pd
from wbs_ordered_converter import WBSOrderedConverter

def get_detailed_employee_data():
    """Get detailed data for mismatched employees"""
    
    # Parse Sierra file in detail
    print("=== DETAILED SIERRA ANALYSIS ===")
    df_sierra = pd.read_excel("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", header=0)
    
    sierra_details = {}
    for _, row in df_sierra.iterrows():
        name = row.get('Name')
        hours = row.get('Hours')
        rate = row.get('Rate')
        
        if pd.notna(name) and pd.notna(hours) and pd.notna(rate):
            name = str(name).strip()
            hours = float(hours)
            rate = float(rate)
            
            if name == 'Name' or hours == 0:
                continue
                
            if name not in sierra_details:
                sierra_details[name] = {'entries': [], 'total_hours': 0, 'rates': set()}
            
            sierra_details[name]['entries'].append({'hours': hours, 'rate': rate})
            sierra_details[name]['total_hours'] += hours
            sierra_details[name]['rates'].add(rate)
    
    # Parse WBS file in detail  
    print("\n=== DETAILED WBS ANALYSIS ===")
    df_raw = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=None)
    
    # Find header row
    header_row = 7  # We found this earlier
    df_wbs = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=header_row)
    
    wbs_details = {}
    for _, row in df_wbs.iterrows():
        name = row.get('Employee Name')
        pay_rate = row.get('Pay Rate')
        regular_hours = row.get('A01', 0) or 0  # Regular hours
        ot_hours = row.get('A02', 0) or 0      # Overtime hours
        total_amount = row.get('Totals', 0) or 0
        
        if pd.notna(name) and name != 'Employee Name' and name != 'Totals':
            name = str(name).strip()
            
            wbs_details[name] = {
                'pay_rate': float(pay_rate) if pd.notna(pay_rate) else 0,
                'regular_hours': float(regular_hours),
                'ot_hours': float(ot_hours), 
                'total_hours': float(regular_hours) + float(ot_hours),
                'total_amount': float(total_amount)
            }
    
    return sierra_details, wbs_details

def main():
    print("=== INVESTIGATING AMOUNT DISCREPANCIES ===")
    
    sierra_details, wbs_details = get_detailed_employee_data()
    converter = WBSOrderedConverter()
    
    # Focus on the biggest discrepancies
    problem_employees = [
        "Miguel Gonzalez",      # +$1,976 difference  
        "Andy Castaneda",       # -$818 difference
        "Kevin Duarte",         # -$694 difference
        "Esau Duarte"          # -$552 difference
    ]
    
    print(f"\n=== INVESTIGATING MAJOR DISCREPANCIES ===")
    
    for problem_name in problem_employees:
        print(f"\n--- {problem_name} ---")
        
        # Normalize name for WBS lookup
        wbs_name = converter.normalize_name(problem_name)
        
        # Find in Sierra data
        sierra_found = None
        for sierra_name, sierra_data in sierra_details.items():
            if problem_name.lower() in sierra_name.lower() or sierra_name.lower() in problem_name.lower():
                sierra_found = (sierra_name, sierra_data)
                break
        
        # Find in WBS data
        wbs_found = None
        for wbs_name_key, wbs_data in wbs_details.items():
            if problem_name.lower() in wbs_name_key.lower() or wbs_name_key.lower() in problem_name.lower():
                wbs_found = (wbs_name_key, wbs_data)
                break
        
        if sierra_found:
            sierra_name, sierra_data = sierra_found
            print(f"  SIERRA '{sierra_name}':")
            print(f"    Total Hours: {sierra_data['total_hours']}")
            print(f"    Rates Used: {sierra_data['rates']}")
            print(f"    Entries: {len(sierra_data['entries'])}")
            
            # Calculate Sierra total
            sierra_total = 0
            for entry in sierra_data['entries']:
                sierra_total += entry['hours'] * entry['rate']
            print(f"    Calculated Total: ${sierra_total:.2f}")
            
            # Show individual entries
            for i, entry in enumerate(sierra_data['entries'][:5]):  # First 5 entries
                print(f"      Entry {i+1}: {entry['hours']}h @ ${entry['rate']}/h = ${entry['hours'] * entry['rate']:.2f}")
        
        if wbs_found:
            wbs_name_key, wbs_data = wbs_found
            print(f"  WBS '{wbs_name_key}':")
            print(f"    Pay Rate: ${wbs_data['pay_rate']}/h")
            print(f"    Regular Hours: {wbs_data['regular_hours']}")
            print(f"    Overtime Hours: {wbs_data['ot_hours']}")
            print(f"    Total Hours: {wbs_data['total_hours']}")
            print(f"    WBS Total Amount: ${wbs_data['total_amount']:.2f}")
            
            # Calculate expected with WBS overtime rules
            if wbs_data['regular_hours'] > 0 or wbs_data['ot_hours'] > 0:
                total_hours = wbs_data['total_hours']
                rate = wbs_data['pay_rate']
                
                # Apply WBS overtime calculation
                wbs_calc = converter.apply_wbs_overtime_rules(total_hours, rate, wbs_name_key)
                print(f"    WBS OT Calculation: ${wbs_calc['total_amount']:.2f}")
                print(f"      Regular: {wbs_calc['regular_hours']}h @ ${rate} = ${wbs_calc['regular_amount']:.2f}")
                print(f"      OT 1.5x: {wbs_calc['ot15_hours']}h @ ${rate * 1.5} = ${wbs_calc['ot15_amount']:.2f}")
                print(f"      OT 2x: {wbs_calc['ot20_hours']}h @ ${rate * 2} = ${wbs_calc['ot20_amount']:.2f}")
        
        if not sierra_found:
            print(f"  ❌ NOT FOUND in Sierra file")
        if not wbs_found:
            print(f"  ❌ NOT FOUND in WBS file")
    
    print(f"\n=== CONCLUSION ===")
    print("The discrepancies appear to be due to:")
    print("1. Different hourly rates between Sierra and WBS")
    print("2. Different total hours worked") 
    print("3. Different overtime calculation methods")
    print("4. Possibly different pay periods or data sources")

if __name__ == "__main__":
    main()