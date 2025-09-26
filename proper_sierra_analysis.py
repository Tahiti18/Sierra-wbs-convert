#!/usr/bin/env python3
"""
Proper analysis of Sierra file to find ALL employees and fix mapping issues
No shortcuts - calculate everything from Sierra data only
"""

import pandas as pd
from wbs_ordered_converter import WBSOrderedConverter

def deep_sierra_analysis():
    """Analyze Sierra file in detail to find all employees"""
    print("=== DEEP SIERRA FILE ANALYSIS ===")
    
    # Read raw Excel data
    df = pd.read_excel("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", header=0)
    
    print(f"Total rows in Sierra file: {len(df)}")
    print(f"Columns: {list(df.columns)}")
    
    # Find all unique names
    all_names = set()
    name_details = {}
    
    for idx, row in df.iterrows():
        name = row.get('Name')
        hours = row.get('Hours')
        rate = row.get('Rate')
        total = row.get('Total')
        
        if pd.notna(name):
            name = str(name).strip()
            
            # Skip header rows and invalid entries
            if name in ['Name', '', 'nan'] or len(name) < 3:
                continue
                
            all_names.add(name)
            
            if name not in name_details:
                name_details[name] = {
                    'entries': [],
                    'total_hours': 0,
                    'total_amount': 0,
                    'rates_used': set(),
                    'row_numbers': []
                }
            
            if pd.notna(hours) and pd.notna(rate):
                hours = float(hours) if hours != 0 else 0
                rate = float(rate) if rate != 0 else 0
                
                if hours > 0 and rate > 0:
                    name_details[name]['entries'].append({
                        'hours': hours,
                        'rate': rate,
                        'amount': hours * rate,
                        'row': idx
                    })
                    name_details[name]['total_hours'] += hours
                    name_details[name]['total_amount'] += hours * rate
                    name_details[name]['rates_used'].add(rate)
                    name_details[name]['row_numbers'].append(idx)
    
    print(f"\nFound {len(all_names)} unique names in Sierra file:")
    
    # Sort by total amount (highest first)
    sorted_employees = []
    for name, details in name_details.items():
        if details['total_hours'] > 0:  # Only employees with actual hours
            sorted_employees.append((name, details))
    
    sorted_employees.sort(key=lambda x: x[1]['total_amount'], reverse=True)
    
    print(f"\nEmployees with payroll data: {len(sorted_employees)}")
    print(f"\n=== TOP 20 SIERRA EMPLOYEES BY AMOUNT ===")
    for i, (name, details) in enumerate(sorted_employees[:20]):
        print(f"{i+1:2d}. {name}: ${details['total_amount']:,.2f} ({details['total_hours']}h, {len(details['rates_used'])} rates)")
    
    print(f"\n=== ALL SIERRA EMPLOYEES ===")
    for name, details in sorted(name_details.items()):
        if details['total_hours'] > 0:
            rates = sorted(details['rates_used'])
            print(f"  {name}: ${details['total_amount']:,.2f} ({details['total_hours']}h @ {rates})")
    
    return name_details

def check_wbs_mapping():
    """Check how Sierra names map to WBS names"""
    print(f"\n=== CHECKING WBS MAPPING ===")
    
    sierra_data = deep_sierra_analysis()
    converter = WBSOrderedConverter()
    
    # Get WBS employee list
    wbs_employees = converter.wbs_order
    print(f"WBS has {len(wbs_employees)} employees")
    
    # Try to map each Sierra employee to WBS
    mapped_count = 0
    unmapped_sierra = []
    
    print(f"\n=== SIERRA TO WBS MAPPING ===")
    for sierra_name, sierra_info in sierra_data.items():
        if sierra_info['total_hours'] > 0:
            # Normalize Sierra name
            wbs_name = converter.normalize_name(sierra_name)
            
            if wbs_name in wbs_employees:
                mapped_count += 1
                print(f"✅ '{sierra_name}' → '{wbs_name}' (${sierra_info['total_amount']:,.2f})")
            else:
                unmapped_sierra.append((sierra_name, wbs_name, sierra_info['total_amount']))
                print(f"❌ '{sierra_name}' → '{wbs_name}' NOT FOUND (${sierra_info['total_amount']:,.2f})")
    
    print(f"\n=== MAPPING RESULTS ===")
    print(f"✅ Successfully mapped: {mapped_count}")
    print(f"❌ Failed to map: {len(unmapped_sierra)}")
    
    if unmapped_sierra:
        print(f"\n=== UNMAPPED SIERRA EMPLOYEES ===")
        total_lost = 0
        for sierra_name, attempted_wbs, amount in unmapped_sierra:
            print(f"  '{sierra_name}' → tried '{attempted_wbs}' (${amount:,.2f})")
            total_lost += amount
        print(f"Total lost payroll: ${total_lost:,.2f}")
    
    # Find WBS employees with no Sierra data
    sierra_mapped_wbs = set()
    for sierra_name, sierra_info in sierra_data.items():
        if sierra_info['total_hours'] > 0:
            wbs_name = converter.normalize_name(sierra_name)
            if wbs_name in wbs_employees:
                sierra_mapped_wbs.add(wbs_name)
    
    missing_from_sierra = []
    for wbs_name in wbs_employees:
        if wbs_name not in sierra_mapped_wbs:
            missing_from_sierra.append(wbs_name)
    
    print(f"\n=== WBS EMPLOYEES MISSING FROM SIERRA ({len(missing_from_sierra)}) ===")
    for wbs_name in missing_from_sierra:
        print(f"  {wbs_name}")
    
    return mapped_count, len(unmapped_sierra), len(missing_from_sierra)

def main():
    print("=== PROPER SIERRA PAYROLL ANALYSIS ===")
    print("Goal: Calculate everything from Sierra data only - no shortcuts!")
    
    mapped, unmapped, missing = check_wbs_mapping()
    
    print(f"\n=== FINAL ANALYSIS ===")
    print(f"Sierra employees mapped to WBS: {mapped}")
    print(f"Sierra employees failed to map: {unmapped}")  
    print(f"WBS employees missing from Sierra: {missing}")
    
    if unmapped > 0:
        print(f"\n🔧 NEXT STEPS:")
        print(f"1. Fix name mappings for {unmapped} unmapped Sierra employees")
        print(f"2. This will recover the lost payroll amounts")
        print(f"3. Only then should we worry about the {missing} missing employees")
    
    if missing > 0:
        print(f"\n⚠️  MISSING EMPLOYEES ISSUE:")
        print(f"These {missing} WBS employees truly don't exist in Sierra file")
        print(f"They should be $0.00 in output (not copied from gold standard)")

if __name__ == "__main__":
    main()