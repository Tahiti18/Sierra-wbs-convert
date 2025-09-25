#!/usr/bin/env python3
"""
Test the FIXED API with correct endpoint
"""
import requests
import json

def test_fixed_api():
    print("=== TESTING FIXED API WITH CORRECT ENDPOINTS ===")
    
    # API base URL 
    base_url = "https://9000-idrlbzy4bg2q93rmh2rr0-6532622b.e2b.dev"
    
    # Test health endpoint
    print("1. Testing health endpoint...")
    health_response = requests.get(f"{base_url}/api/health")
    print(f"Health Status: {health_response.status_code}")
    if health_response.status_code == 200:
        health_data = health_response.json()
        print(f"Version: {health_data.get('version', 'unknown')}")
        print(f"Converter: {health_data.get('converter', 'unknown')}")
        print(f"Employee DB Count: {health_data.get('employee_database_count', 'unknown')}")
    
    # Test file upload and conversion using correct endpoint
    print(f"\n2. Testing file upload and conversion via /api/process-payroll...")
    
    try:
        # Upload Sierra file
        with open('/home/user/webapp/sierra_input_new.xlsx', 'rb') as file:
            files = {
                'sierra_file': file  # The endpoint expects 'sierra_file'
            }
            
            upload_response = requests.post(f"{base_url}/api/process-payroll", files=files)
            print(f"Upload Status: {upload_response.status_code}")
            
            if upload_response.status_code == 200:
                # Check if response is JSON or file
                content_type = upload_response.headers.get('content-type', '')
                
                if 'application/json' in content_type:
                    result = upload_response.json()
                    print(f"JSON Response: {json.dumps(result, indent=2)}")
                    
                elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type or 'application/octet-stream' in content_type:
                    # It's an Excel file - save it
                    with open('/home/user/webapp/api_converted_output.xlsx', 'wb') as f:
                        f.write(upload_response.content)
                    print(f"✅ Received Excel file, saved as: api_converted_output.xlsx")
                    
                    # Quick verification of the output
                    print(f"\n3. Quick verification of API output...")
                    import pandas as pd
                    
                    try:
                        df = pd.read_excel('/home/user/webapp/api_converted_output.xlsx')
                        print(f"   Output shape: {df.shape}")
                        
                        # Look for Dianne
                        dianne_rows = df[df.iloc[:, 2].astype(str).str.contains("Dianne", na=False)]
                        if not dianne_rows.empty:
                            dianne_data = dianne_rows.iloc[0]
                            print(f"   Dianne found:")
                            print(f"     Employee#: {dianne_data.iloc[0]}")
                            print(f"     SSN: {dianne_data.iloc[1]}")
                            print(f"     Name: {dianne_data.iloc[2]}")
                            print(f"     Rate: {dianne_data.iloc[5]}")
                            print(f"     Hours: {dianne_data.iloc[7]}")
                            print(f"     Total: {dianne_data.iloc[27]}")
                            
                            # Verify it matches expected values
                            expected = {"emp": "0000662082", "ssn": "626946016", "total": 112.0}
                            actual_emp = str(int(dianne_data.iloc[0])) if dianne_data.iloc[0] else "0"
                            actual_ssn = str(int(dianne_data.iloc[1])) if dianne_data.iloc[1] else "0"
                            actual_total = float(dianne_data.iloc[27]) if dianne_data.iloc[27] else 0
                            
                            emp_match = actual_emp.endswith("662082")  # Employee number should end with this
                            ssn_match = actual_ssn == expected["ssn"]
                            total_match = abs(actual_total - expected["total"]) < 0.01
                            
                            print(f"     ✅ Verification: Employee={'✅' if emp_match else '❌'}, SSN={'✅' if ssn_match else '❌'}, Total={'✅' if total_match else '❌'}")
                            
                            if emp_match and ssn_match and total_match:
                                print(f"   🎉 API OUTPUT MATCHES GOLD STANDARD!")
                            
                        else:
                            print(f"   ⚠️  Dianne not found in API output")
                            
                    except Exception as e:
                        print(f"   ❌ Error verifying output: {e}")
                        
                else:
                    print(f"Unexpected content type: {content_type}")
                    print(f"Response text: {upload_response.text[:500]}...")
                
            else:
                print(f"❌ Conversion failed: {upload_response.status_code}")
                print(f"Error: {upload_response.text}")
                
    except Exception as e:
        print(f"❌ Test failed: {e}")

    # Test the multi-stage endpoint as well
    print(f"\n4. Testing multi-stage processing...")
    
    try:
        with open('/home/user/webapp/sierra_input_new.xlsx', 'rb') as file:
            files = {
                'sierra_file': file
            }
            
            multi_response = requests.post(f"{base_url}/api/multi-stage/process-all", files=files)
            print(f"Multi-stage Status: {multi_response.status_code}")
            
            if multi_response.status_code == 200:
                result = multi_response.json()
                print(f"Multi-stage Success: {result.get('success', False)}")
                print(f"Stages completed: {len(result.get('stages', []))}")
                if 'final_result' in result:
                    print(f"Final result keys: {list(result['final_result'].keys())}")
                    
    except Exception as e:
        print(f"❌ Multi-stage test failed: {e}")

if __name__ == "__main__":
    test_fixed_api()