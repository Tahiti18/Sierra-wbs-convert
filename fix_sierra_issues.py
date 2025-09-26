#!/usr/bin/env python3
"""
Fix the real Sierra mapping and calculation issues
Calculate everything properly from Sierra data
"""

from wbs_ordered_converter import WBSOrderedConverter
import pandas as pd

def fix_name_mappings():
    """Fix the name mapping issues"""
    print("=== FIXING NAME MAPPING ISSUES ===")
    
    # The issues found:
    # 1. "Daniel Carrasco" doesn't map to any WBS employee
    # 2. Both Kevin Cortez and Kevin Duarte map to "Duarte, Kevin"
    
    print("Issues to fix:")
    print("1. Daniel Carrasco ($2,200) - needs proper WBS mapping")
    print("2. Kevin Cortez ($784) vs Kevin Duarte ($1,684) - both mapping to 'Duarte, Kevin'")
    
    # Need to check which WBS employees might match these
    converter = WBSOrderedConverter()
    wbs_employees = converter.wbs_order
    
    print(f"\nLooking for potential matches in WBS list...")
    
    # Look for Daniel/Carrasco matches
    daniel_matches = [name for name in wbs_employees if 'daniel' in name.lower() or 'carrasco' in name.lower()]
    print(f"Potential Daniel Carrasco matches: {daniel_matches}")
    
    # Look for Kevin matches  
    kevin_matches = [name for name in wbs_employees if 'kevin' in name.lower()]
    print(f"Potential Kevin matches: {kevin_matches}")
    
    # The solution should be:
    # - Kevin Cortez → Cortez, Kevin (if exists)
    # - Kevin Duarte → Duarte, Kevin (keep existing) 
    # - Daniel Carrasco → find appropriate match or create mapping
    
    return True

def analyze_calculation_differences():
    """Analyze why Sierra calculations differ from WBS gold standard"""
    print(f"\n=== ANALYZING CALCULATION DIFFERENCES ===")
    
    # Get Sierra data
    df_sierra = pd.read_excel("Sierra_Payroll_9_12_25_for_Marwan_2.xlsx", header=0)
    
    # Parse WBS gold standard 
    df_wbs = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=7)
    
    # Focus on employees with biggest discrepancies
    problem_cases = [
        ("Miguel Gonzalez", "Gonzalez, Miguel"),  # Sierra: $2,296 vs WBS: $2,440
        ("Kevin Duarte", "Duarte, Kevin"),        # Sierra: $1,684 vs WBS: $1,828  
        ("Efrain Santos", "Santos, Efrain"),      # Sierra: $1,372 vs WBS: $1,516
        ("Javier Santos", "Santos, Javier")       # Sierra: $1,308 vs WBS: $1,452
    ]
    
    print(f"Analyzing {len(problem_cases)} major discrepancies...")
    
    for sierra_name, wbs_name in problem_cases:
        print(f"\n--- {sierra_name} → {wbs_name} ---")
        
        # Get Sierra details
        sierra_entries = []
        sierra_total = 0
        
        for _, row in df_sierra.iterrows():
            name = row.get('Name')
            if pd.notna(name) and str(name).strip() == sierra_name:
                hours = row.get('Hours', 0)
                rate = row.get('Rate', 0)
                if pd.notna(hours) and pd.notna(rate) and hours > 0:
                    hours = float(hours)
                    rate = float(rate)
                    amount = hours * rate
                    sierra_entries.append((hours, rate, amount))
                    sierra_total += amount
        
        print(f"  Sierra '{sierra_name}': ${sierra_total:.2f}")
        for i, (h, r, a) in enumerate(sierra_entries[:3]):  # First 3 entries
            print(f"    Entry {i+1}: {h}h @ ${r}/h = ${a:.2f}")
        if len(sierra_entries) > 3:
            print(f"    ... and {len(sierra_entries) - 3} more entries")
        
        # Get WBS details
        wbs_info = None
        for _, row in df_wbs.iterrows():
            name = row.get('Employee Name')
            if pd.notna(name) and str(name).strip() == wbs_name:
                wbs_info = {
                    'pay_rate': row.get('Pay Rate', 0),
                    'regular_hours': row.get('A01', 0) or 0,
                    'ot_hours': row.get('A02', 0) or 0,
                    'total_amount': row.get('Totals', 0) or 0
                }
                break
        
        if wbs_info:
            print(f"  WBS '{wbs_name}': ${wbs_info['total_amount']:.2f}")
            print(f"    Rate: ${wbs_info['pay_rate']}/h")
            print(f"    Regular: {wbs_info['regular_hours']}h")  
            print(f"    Overtime: {wbs_info['ot_hours']}h")
            print(f"    Total Hours: {wbs_info['regular_hours'] + wbs_info['ot_hours']}h")
            
            # The issue: Sierra has multiple rates, WBS has single rate + overtime calculation
            # Sierra should be consolidated to match WBS overtime approach

def create_proper_sierra_converter():
    """Create a proper Sierra converter that calculates everything correctly"""
    print(f"\n=== CREATING PROPER SIERRA CONVERTER ===")
    
    print("Strategy:")
    print("1. Fix the 1 unmapped employee (Daniel Carrasco)")
    print("2. Fix duplicate Kevin mapping issue") 
    print("3. Consolidate Sierra hours per employee properly")
    print("4. Apply WBS overtime rules to consolidated hours")
    print("5. Set missing employees to $0.00 (not copy from gold standard)")
    
    # Key fixes needed in wbs_ordered_converter.py:
    fixes_needed = [
        "Add mapping for Daniel Carrasco to appropriate WBS employee",
        "Fix Kevin Cortez vs Kevin Duarte mapping conflict",
        "Ensure proper hour consolidation per employee",
        "Apply correct WBS overtime calculations",
        "Handle missing employees as $0.00"
    ]
    
    print(f"\nFixes needed in converter:")
    for i, fix in enumerate(fixes_needed, 1):
        print(f"  {i}. {fix}")
    
    return fixes_needed

def main():
    print("=== FIXING SIERRA CONVERSION ISSUES PROPERLY ===")
    print("Goal: Calculate everything from Sierra data - NO shortcuts!")
    
    fix_name_mappings()
    analyze_calculation_differences() 
    fixes_needed = create_proper_sierra_converter()
    
    print(f"\n=== SUMMARY OF REQUIRED FIXES ===")
    print(f"Current status:")
    print(f"  ✅ 65 Sierra employees mapping correctly") 
    print(f"  ❌ 1 Sierra employee not mapping ($2,200 lost)")
    print(f"  ⚠️  16 WBS employees missing from Sierra (should be $0.00)")
    print(f"  🔧 Multiple calculation discrepancies to fix")
    
    print(f"\nNext steps:")
    print(f"1. Fix name mappings to recover $2,200")
    print(f"2. Fix calculation method for proper amounts") 
    print(f"3. Test with Sierra data only")
    print(f"4. Missing WBS employees = $0.00 (correct behavior)")

if __name__ == "__main__":
    main()