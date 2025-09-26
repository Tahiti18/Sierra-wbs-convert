#!/usr/bin/env python3
"""
Fix name matching between Sierra and WBS files
"""

import pandas as pd
from wbs_ordered_converter import WBSOrderedConverter

def get_all_sierra_names():
    """Get all unique names from Sierra file"""
    df = pd.read_excel("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", header=0)
    
    names = set()
    for _, row in df.iterrows():
        name = row.get('Name')
        if pd.notna(name):
            name = str(name).strip()
            if name != 'Name' and len(name) > 3:
                names.add(name)
    
    return sorted(names)

def get_all_wbs_names():
    """Get all names from WBS file"""
    df = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=7)
    
    names = set()
    for _, row in df.iterrows():
        name = row.get('Employee Name')
        if pd.notna(name):
            name = str(name).strip()
            if name != 'Employee Name' and name != 'Totals' and len(name) > 3:
                names.add(name)
    
    return sorted(names)

def create_name_mapping():
    """Create mapping between Sierra and WBS names"""
    
    sierra_names = get_all_sierra_names()
    wbs_names = get_all_wbs_names()
    converter = WBSOrderedConverter()
    
    print(f"Sierra file has {len(sierra_names)} unique names")
    print(f"WBS file has {len(wbs_names)} unique names")
    
    # Try to match Sierra names to WBS names
    matches = {}
    sierra_normalized = {}
    
    for sierra_name in sierra_names:
        # Normalize to "Last, First" format
        normalized = converter.normalize_name(sierra_name)
        sierra_normalized[sierra_name] = normalized
        
        # Check if normalized name exists in WBS
        if normalized in wbs_names:
            matches[sierra_name] = normalized
    
    print(f"\n=== NAME MATCHING RESULTS ===")
    print(f"Direct matches: {len(matches)}")
    
    # Show matches
    print(f"\n=== SUCCESSFUL MATCHES ===")
    for sierra, wbs in sorted(matches.items()):
        print(f"  '{sierra}' → '{wbs}'")
    
    # Show unmatched Sierra names
    unmatched_sierra = []
    for sierra_name in sierra_names:
        if sierra_name not in matches:
            normalized = sierra_normalized[sierra_name]
            unmatched_sierra.append((sierra_name, normalized))
    
    print(f"\n=== UNMATCHED SIERRA NAMES ({len(unmatched_sierra)}) ===")
    for sierra, normalized in unmatched_sierra:
        print(f"  '{sierra}' → normalized: '{normalized}'")
    
    # Show unmatched WBS names
    matched_wbs = set(matches.values())
    unmatched_wbs = [name for name in wbs_names if name not in matched_wbs]
    
    print(f"\n=== UNMATCHED WBS NAMES ({len(unmatched_wbs)}) ===")
    for wbs_name in unmatched_wbs:
        print(f"  '{wbs_name}'")
    
    # Try fuzzy matching for unmatched names
    print(f"\n=== ATTEMPTING FUZZY MATCHING ===")
    potential_matches = []
    
    for sierra_name, normalized in unmatched_sierra:
        sierra_parts = sierra_name.lower().split()
        
        for wbs_name in unmatched_wbs:
            wbs_parts = wbs_name.lower().split()
            
            # Check if first and last names match in any order
            sierra_first = sierra_parts[0] if sierra_parts else ""
            sierra_last = sierra_parts[-1] if len(sierra_parts) > 1 else ""
            
            wbs_last = wbs_parts[0].replace(",", "") if wbs_parts else ""
            wbs_first = wbs_parts[1] if len(wbs_parts) > 1 else ""
            
            if (sierra_first == wbs_first and sierra_last == wbs_last) or \
               (sierra_first in wbs_name.lower() and sierra_last in wbs_name.lower()):
                potential_matches.append((sierra_name, wbs_name))
    
    print(f"Found {len(potential_matches)} potential fuzzy matches:")
    for sierra, wbs in potential_matches:
        print(f"  '{sierra}' ≈ '{wbs}'")
    
    return matches, potential_matches

def main():
    print("=== FIXING NAME MATCHING BETWEEN SIERRA AND WBS ===")
    
    matches, potential_matches = create_name_mapping()
    
    print(f"\n=== SUMMARY ===")
    print(f"Direct matches: {len(matches)}")
    print(f"Potential matches: {len(potential_matches)}")
    print(f"Total possible matches: {len(matches) + len(potential_matches)}")
    
    # Now let's see what the match rate would be
    sierra_names = get_all_sierra_names()
    wbs_names = get_all_wbs_names()
    
    coverage = (len(matches) + len(potential_matches)) / len(sierra_names) * 100
    print(f"Sierra name coverage: {coverage:.1f}%")

if __name__ == "__main__":
    main()