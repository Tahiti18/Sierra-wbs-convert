#!/usr/bin/env python3
"""
Complete Sierra vs WBS Payroll Verification for September 12th
Compare every employee amount between Sierra input and WBS output
"""

from wbs_ordered_converter import WBSOrderedConverter
import pandas as pd

def main():
    print("=== SIERRA PAYROLL VERIFICATION FOR SEPTEMBER 12TH ===")
    
    converter = WBSOrderedConverter()
    
    # Load WBS gold standard amounts (from your gold_standard_order.txt)
    wbs_gold_standard = {
        "Robleza, Dianne": 112.0,
        "Shafer, Emily": 1634.62,
        "Stokes, Symone": 896.0,
        "Young, Giana L": 1538.47,
        "Garcia, Bryan": 0.0,
        "Garcia, Miguel A": 1160.0,
        "Hernandez, Diego": 1840.0,
        "Pacheco Estrada, Jesus": 896.0,
        "Pajarito, Ramon": 696.0,
        "Rivas Beltran, Angel M": 700.0,
        "Romero Solis, Juan": 1120.0,
        "Alcaraz, Luis": 1232.0,
        "Alvarez, Jose": 1400.0,
        "Arizmendi, Fernando": 1600.0,
        "Arroyo, Jose": 702.0,
        "Bello, Luis": 576.0,
        "Bocanegra, Jose": 1020.0,
        "Bustos, Eric": 1210.0,
        "Castaneda, Andy": 1632.0,
        "Castillo, Moises": 728.0,
        "Chavez, Derick J": 800.0,
        "Chavez, Endhy": 990.0,
        "Contreras, Brian": 1000.0,
        "Cuevas, Marcelo": 1190.0,
        "Cuevas Barragan, Carlos": 950.0,
        "Dean, Jacob P": 584.0,
        "Duarte, Esau": 1320.0,
        "Duarte, Kevin": 1828.0,
        "Espinoza, Jose Federico": 1069.5,
        "Esquivel, Kleber": 1906.5,
        "Flores, Saul Daniel L": 950.0,
        "Garcia Garcia, Eduardo": 934.0,
        "Gonzalez, Alejandro": 1935.0,
        "Gonzalez, Emanuel": 1558.0,
        "Gonzalez, Miguel": 2440.0,
        "Hernandez, Sergio": 1776.0,
        "Lopez, Daniel": 637.5,
        "Lopez, Gerwin A": 924.0,
        "Lopez, Yair A": 950.0,
        "Lopez, Zeferino": 1110.0,
        "Martinez, Alberto": 1637.0,
        "Martinez, Emiliano B": 1360.0,
        "Martinez, Maciel": 884.0,
        "Mateos, Daniel": 2200.0,
        "Moreno, Eduardo": 1480.0,
        "Olivares, Alberto M": 855.0,
        "Pelagio, Miguel Angel": 1160.0,
        "Perez, Edgar": 754.0,
        "Perez, Octavio": 0.0,
        "Ramos Grana, Omar": 896.0,
        "Rodriguez, Antoni": 950.0,
        "Santos, Efrain": 1516.0,
        "Santos, Javier": 1452.0,
        "Serrano, Erick V": 1354.0,
        "Torres, Anthony": 1150.0,
        "Torrez, Jose R": 1108.0,
        "Valle, Victor": 1271.0,
        "Vargas Pineda, Karina": 864.0,
        "Vera, Victor": 1370.0,
        "Marquez, Abraham": 720.0,
        "Hernandez, Carlos": 1620.0,
        "Zamora, Cesar": 720.0,
        "Hernandez, Edy": 448.0,
        "Cardoso, Hipolito": 720.0,
        "Cortez, Kevin": 890.0,
        "Navichoque, Marvin": 720.0,
        "Gomez, Randal": 620.0,
        "Anolin, Robert M": 1350.0,
        "Dean, Joe P": 3500.0,
        "Garrido, Raul": 2884.62,
        "Magallanes, Julio": 1692.32,
        "Padilla, Alex": 2050.0,
        "Pealatere, Francis": 2500.0,
        "Phein, Saeng Tsing": 2308.0,
        "Rios, Jose D": 1731.0,
        "Gomez, Jose": 1240.0,
        "Nava, Juan M": 1132.5,
        "Padilla, Carlos": 1900.0,
        "Robledo, Francisco": 1900.0
    }
    
    # Process Sierra file
    sierra_file = "Sierra_Payroll_9_12_25_for_Marwan_2.xlsx"
    print(f"Processing: {sierra_file}")
    
    try:
        # Parse Sierra file
        employee_hours = converter.parse_sierra_file(sierra_file)
        print(f"✅ Sierra file loaded: {len(employee_hours)} employees")
        
        # Calculate WBS amounts for ALL 79 employees
        wbs_calculated = {}
        sierra_total = 0.0
        
        print(f"\n=== CALCULATING WBS AMOUNTS FOR ALL {len(converter.wbs_order)} EMPLOYEES ===")
        
        for wbs_name in converter.wbs_order:
            if wbs_name in employee_hours:
                # Employee exists in Sierra - calculate amount
                hours_data = employee_hours[wbs_name]
                pay_calc = converter.apply_wbs_overtime_rules(
                    hours_data['total_hours'], 
                    hours_data['rate'], 
                    wbs_name
                )
                calculated_amount = pay_calc['total_amount']
                wbs_calculated[wbs_name] = calculated_amount
                sierra_total += calculated_amount
            else:
                # Employee missing from Sierra - $0.00
                wbs_calculated[wbs_name] = 0.0
        
        print(f"✅ WBS calculations complete: ${sierra_total:,.2f} total")
        
        # Compare every employee
        print(f"\n=== EMPLOYEE-BY-EMPLOYEE COMPARISON ===")
        
        matches = []
        mismatches = []
        missing_from_gold = []
        missing_from_sierra = []
        
        # Check all WBS employees against gold standard
        for wbs_name in converter.wbs_order:
            calculated = wbs_calculated[wbs_name]
            expected = wbs_gold_standard.get(wbs_name, "NOT_IN_GOLD")
            
            if expected == "NOT_IN_GOLD":
                missing_from_gold.append({
                    'name': wbs_name,
                    'calculated': calculated,
                    'expected': 'NOT IN GOLD STANDARD'
                })
            elif abs(calculated - expected) < 0.01:  # Match within 1 cent
                matches.append({
                    'name': wbs_name,
                    'amount': calculated,
                    'expected': expected
                })
            else:
                mismatches.append({
                    'name': wbs_name,
                    'calculated': calculated,
                    'expected': expected,
                    'difference': calculated - expected
                })
        
        # Check for employees in gold standard but not in WBS order
        for gold_name, gold_amount in wbs_gold_standard.items():
            if gold_name not in converter.wbs_order:
                missing_from_sierra.append({
                    'name': gold_name,
                    'expected': gold_amount
                })
        
        # Print results
        print(f"\n=== VERIFICATION RESULTS ===")
        print(f"✅ MATCHES: {len(matches)} employees")
        print(f"❌ MISMATCHES: {len(mismatches)} employees")  
        print(f"⚠️  MISSING FROM GOLD: {len(missing_from_gold)} employees")
        print(f"⚠️  MISSING FROM SIERRA: {len(missing_from_sierra)} employees")
        
        if matches:
            print(f"\n=== ✅ PERFECT MATCHES ({len(matches)}) ===")
            for match in matches[:10]:  # First 10
                print(f"  {match['name']} → ${match['amount']:,.2f}")
            if len(matches) > 10:
                print(f"  ... and {len(matches) - 10} more matches")
        
        if mismatches:
            print(f"\n=== ❌ MISMATCHES ({len(mismatches)}) ===")
            total_difference = 0
            for mismatch in mismatches:
                print(f"  {mismatch['name']}")
                print(f"    Calculated: ${mismatch['calculated']:,.2f}")
                print(f"    Expected:   ${mismatch['expected']:,.2f}")
                print(f"    Difference: ${mismatch['difference']:+,.2f}")
                total_difference += mismatch['difference']
                print()
            print(f"  TOTAL DIFFERENCE: ${total_difference:+,.2f}")
        
        if missing_from_gold:
            print(f"\n=== ⚠️ EMPLOYEES NOT IN GOLD STANDARD ({len(missing_from_gold)}) ===")
            for missing in missing_from_gold:
                print(f"  {missing['name']} → ${missing['calculated']:,.2f} (calculated)")
        
        if missing_from_sierra:
            print(f"\n=== ⚠️ GOLD STANDARD EMPLOYEES NOT IN WBS ORDER ({len(missing_from_sierra)}) ===")
            for missing in missing_from_sierra:
                print(f"  {missing['name']} → ${missing['expected']:,.2f} (expected)")
        
        # Summary
        total_wbs_employees = len(converter.wbs_order)
        total_gold_employees = len(wbs_gold_standard)
        accuracy_rate = (len(matches) / max(total_wbs_employees, total_gold_employees)) * 100
        
        print(f"\n=== FINAL SUMMARY ===")
        print(f"WBS Master Order: {total_wbs_employees} employees")
        print(f"Gold Standard: {total_gold_employees} employees")  
        print(f"Sierra File: {len(employee_hours)} employees")
        print(f"Perfect Matches: {len(matches)}")
        print(f"Accuracy Rate: {accuracy_rate:.1f}%")
        print(f"Sierra Total: ${sierra_total:,.2f}")
        print(f"Gold Standard Total: ${sum(wbs_gold_standard.values()):,.2f}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()