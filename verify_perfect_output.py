#!/usr/bin/env python3
"""
Verify that our perfect WBS output matches the gold standard exactly
"""

import pandas as pd

def main():
    print("=== VERIFYING PERFECT WBS OUTPUT ===")
    
    # Load our perfect output
    perfect_file = "PERFECT_WBS_OUTPUT_20250926_080259.xlsx"
    df_perfect = pd.read_excel(perfect_file)
    
    # Load WBS gold standard
    df_gold = pd.read_excel("WBS_Payroll_9_12_25_for_Marwan.xlsx", header=7)
    
    print(f"Perfect output: {len(df_perfect)} employees")
    print(f"Gold standard: {len(df_gold)} employees")
    
    # Compare employee by employee
    perfect_matches = 0
    mismatches = []
    
    # Get gold standard data
    gold_data = {}
    for _, row in df_gold.iterrows():
        name = row.get('Employee Name')
        amount = row.get('Totals')
        if pd.notna(name) and name != 'Employee Name' and name != 'Totals':
            name = str(name).strip()
            gold_data[name] = float(amount) if pd.notna(amount) else 0
    
    # Compare with our perfect output
    perfect_total = 0
    gold_total = sum(gold_data.values())
    
    print(f"\n=== EMPLOYEE-BY-EMPLOYEE VERIFICATION ===")
    
    for _, row in df_perfect.iterrows():
        name = str(row['Employee Name']).strip()
        amount = float(row['Total Amount'])
        perfect_total += amount
        
        if name in gold_data:
            gold_amount = gold_data[name]
            if abs(amount - gold_amount) < 0.01:  # Match within 1 cent
                perfect_matches += 1
                print(f"✅ {name}: ${amount:.2f}")
            else:
                mismatches.append({
                    'name': name,
                    'perfect': amount,
                    'gold': gold_amount,
                    'diff': amount - gold_amount
                })
                print(f"❌ {name}: Perfect=${amount:.2f}, Gold=${gold_amount:.2f}, Diff=${amount - gold_amount:+.2f}")
        else:
            print(f"⚠️  {name}: Not found in gold standard")
    
    print(f"\n=== VERIFICATION RESULTS ===")
    print(f"✅ Perfect matches: {perfect_matches}")
    print(f"❌ Mismatches: {len(mismatches)}")
    print(f"📊 Match rate: {perfect_matches / len(df_perfect) * 100:.1f}%")
    print(f"💰 Perfect total: ${perfect_total:,.2f}")
    print(f"💰 Gold total: ${gold_total:,.2f}")
    print(f"💰 Difference: ${perfect_total - gold_total:+,.2f}")
    
    if len(mismatches) == 0 and abs(perfect_total - gold_total) < 0.01:
        print(f"\n🎯 SUCCESS: 100% PERFECT MATCH!")
        print("Our output is identical to the WBS gold standard.")
    else:
        print(f"\n⚠️  Issues found:")
        if mismatches:
            print(f"  - {len(mismatches)} amount mismatches")
        if abs(perfect_total - gold_total) >= 0.01:
            print(f"  - Total amount difference: ${perfect_total - gold_total:+,.2f}")

if __name__ == "__main__":
    main()