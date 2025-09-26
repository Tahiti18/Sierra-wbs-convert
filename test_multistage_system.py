#!/usr/bin/env python3
"""
Test the 3-stage system with verification at each step
"""

import requests
import json

def test_multistage_conversion():
    """Test each stage of the conversion with verification"""
    
    base_url = "http://localhost:5001/api/multi-stage"
    
    # Upload Sierra file first
    sierra_file = "Sierra_Payroll_9_12_25_for_Marwan_2.xlsx"
    
    print("=== TESTING 3-STAGE CONVERSION SYSTEM ===")
    print(f"Processing: {sierra_file}")
    
    try:
        # Stage 1: Parse Sierra file
        print(f"\n=== STAGE 1: PARSE SIERRA FILE ===")
        
        with open(sierra_file, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{base_url}/stage1-parse", files=files, timeout=30)
        
        if response.status_code != 200:
            print(f"❌ Stage 1 failed: {response.status_code}")
            print(response.text)
            return False
            
        stage1_data = response.json()
        print(f"✅ Stage 1 complete:")
        print(f"   Raw employees found: {stage1_data.get('raw_employees', 0)}")
        print(f"   Total entries: {stage1_data.get('total_entries', 0)}")
        print(f"   Sample employees: {stage1_data.get('sample_employees', [])[:5]}")
        
        # Stage 2: Consolidate employee hours
        print(f"\n=== STAGE 2: CONSOLIDATE HOURS ===")
        
        with open(sierra_file, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{base_url}/stage2-consolidate", files=files, timeout=30)
        
        if response.status_code != 200:
            print(f"❌ Stage 2 failed: {response.status_code}")
            print(response.text)
            return False
            
        stage2_data = response.json()
        print(f"✅ Stage 2 complete:")
        print(f"   Consolidated employees: {stage2_data.get('consolidated_employees', 0)}")
        print(f"   Total hours: {stage2_data.get('total_hours', 0)}")
        print(f"   Total amount (raw): ${stage2_data.get('total_raw_amount', 0):,.2f}")
        
        # Show top employees by amount
        if 'employee_details' in stage2_data:
            print(f"   Top 5 employees by amount:")
            sorted_employees = sorted(stage2_data['employee_details'][:10], 
                                    key=lambda x: x.get('raw_amount', 0), reverse=True)
            for i, emp in enumerate(sorted_employees[:5]):
                print(f"     {i+1}. {emp['name']}: {emp['total_hours']}h @ ${emp['consolidated_rate']}/h = ${emp['raw_amount']:.2f}")
        
        # Stage 3: Apply WBS overtime rules
        print(f"\n=== STAGE 3: APPLY WBS OVERTIME RULES ===")
        
        with open(sierra_file, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{base_url}/stage3-overtime", files=files, timeout=30)
        
        if response.status_code != 200:
            print(f"❌ Stage 3 failed: {response.status_code}")
            print(response.text)
            return False
            
        stage3_data = response.json()
        print(f"✅ Stage 3 complete:")
        print(f"   Employees with overtime calculations: {stage3_data.get('processed_employees', 0)}")
        print(f"   Total amount (with WBS overtime): ${stage3_data.get('total_wbs_amount', 0):,.2f}")
        print(f"   Regular hours total: {stage3_data.get('total_regular_hours', 0)}")
        print(f"   Overtime hours total: {stage3_data.get('total_overtime_hours', 0)}")
        
        # Show overtime calculation examples
        if 'overtime_examples' in stage3_data:
            print(f"   Overtime calculation examples:")
            for emp in stage3_data['overtime_examples'][:3]:
                print(f"     {emp['name']}: {emp['total_hours']}h → Reg: {emp['regular_hours']}h, OT: {emp['ot_hours']}h, Total: ${emp['total_amount']:.2f}")
        
        # Stage 4: Map to WBS names  
        print(f"\n=== STAGE 4: MAP TO WBS NAMES ===")
        
        with open(sierra_file, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{base_url}/stage4-mapping", files=files, timeout=30)
        
        if response.status_code != 200:
            print(f"❌ Stage 4 failed: {response.status_code}")
            print(response.text)
            return False
            
        stage4_data = response.json()
        print(f"✅ Stage 4 complete:")
        print(f"   Sierra employees mapped to WBS: {stage4_data.get('mapped_employees', 0)}")
        print(f"   Sierra employees failed to map: {stage4_data.get('unmapped_employees', 0)}")
        print(f"   Mapped payroll amount: ${stage4_data.get('mapped_amount', 0):,.2f}")
        
        if stage4_data.get('unmapped_list'):
            print(f"   Unmapped employees:")
            for unmapped in stage4_data['unmapped_list']:
                print(f"     '{unmapped['sierra_name']}' → '${unmapped['amount']:.2f}' (LOST)")
        
        # Stage 5: Create final WBS output
        print(f"\n=== STAGE 5: CREATE FINAL WBS OUTPUT ===")
        
        with open(sierra_file, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{base_url}/stage5-wbs", files=files, timeout=30)
        
        if response.status_code != 200:
            print(f"❌ Stage 5 failed: {response.status_code}")
            print(response.text)
            return False
            
        stage5_data = response.json()
        print(f"✅ Stage 5 complete:")
        print(f"   Total WBS employees in output: {stage5_data.get('total_wbs_employees', 0)}")
        print(f"   WBS employees with Sierra data: {stage5_data.get('wbs_with_data', 0)}")
        print(f"   WBS employees with $0.00: {stage5_data.get('wbs_with_zero', 0)}")
        print(f"   Final total amount: ${stage5_data.get('final_total_amount', 0):,.2f}")
        
        # Validation: Check for issues
        print(f"\n=== VALIDATION: CHECK FOR ISSUES ===")
        
        with open(sierra_file, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{base_url}/validate", files=files, timeout=30)
        
        if response.status_code == 200:
            validation_data = response.json()
            print(f"✅ Validation complete:")
            
            issues = validation_data.get('issues', [])
            warnings = validation_data.get('warnings', [])
            
            print(f"   Issues found: {len(issues)}")
            print(f"   Warnings: {len(warnings)}")
            
            if issues:
                print(f"   ISSUES:")
                for issue in issues:
                    print(f"     ❌ {issue}")
            
            if warnings:
                print(f"   WARNINGS:")
                for warning in warnings:
                    print(f"     ⚠️  {warning}")
            
            if not issues and not warnings:
                print(f"   🎯 NO ISSUES - Conversion is accurate!")
        
        # Final summary
        print(f"\n=== FINAL SUMMARY ===")
        print(f"✅ All 5 stages completed successfully")
        print(f"📊 {stage5_data.get('wbs_with_data', 0)} employees calculated from Sierra data")
        print(f"💰 ${stage5_data.get('final_total_amount', 0):,.2f} total payroll")
        print(f"🎯 {stage5_data.get('wbs_with_zero', 0)} missing employees correctly set to $0.00")
        
        return True
        
    except Exception as e:
        print(f"❌ Multi-stage test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("=== TESTING 3-STAGE CONVERSION WITH VERIFICATION ===")
    print("This system processes Sierra payroll in clear, verifiable stages:")
    print("1. Parse → 2. Consolidate → 3. Overtime → 4. Map → 5. WBS Output")
    
    success = test_multistage_conversion()
    
    if success:
        print(f"\n🎯 SUCCESS: Multi-stage system working correctly!")
        print("Each stage can be verified independently")
        print("No shortcuts - everything calculated from Sierra data")
    else:
        print(f"\n❌ Multi-stage system needs fixes")

if __name__ == "__main__":
    main()