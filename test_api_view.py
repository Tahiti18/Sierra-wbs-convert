#!/usr/bin/env python3
"""
Test the API view mode functionality directly
"""

import requests
import json

def test_view_mode():
    print("=== TESTING VIEW MODE API ===")
    
    # API endpoint
    api_url = "https://5001-idrlbzy4bg2q93rmh2rr0-6532622b.e2b.dev/api/process-payroll"
    
    # Test file
    sierra_file = "Sierra_Payroll_9_12_25_for_Marwan_2.xlsx"
    
    try:
        # Prepare the request
        with open(sierra_file, 'rb') as f:
            files = {'file': f}
            data = {'format': 'json'}  # Request JSON view mode
            
            print(f"Sending VIEW mode request with {sierra_file}...")
            response = requests.post(api_url, files=files, data=data, timeout=60)
            
            if response.status_code == 200:
                result = response.json()
                print("✅ VIEW mode API successful!")
                print(f"Total employees: {result.get('total_employees', 'N/A')}")
                print(f"Sierra employees: {result.get('sierra_employees', 'N/A')}")
                print(f"Total amount: ${result.get('summary', {}).get('total_amount', 0):,.2f}")
                
                # Check first 5 employees
                if 'full_wbs_data' in result:
                    print("\nFirst 5 WBS employees in response:")
                    for i, emp in enumerate(result['full_wbs_data'][:5]):
                        print(f"   {i+1}. {emp.get('ssn', 'N/A')} - {emp.get('employee_name', 'N/A')} - ${emp.get('total_amount', 0):,.2f}")
                        
                    print(f"\nTotal employees returned: {len(result['full_wbs_data'])}")
                else:
                    print("❌ No full_wbs_data in response")
            else:
                print(f"❌ API error: {response.status_code}")
                print(response.text)
                
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    test_view_mode()