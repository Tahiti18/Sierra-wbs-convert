#!/usr/bin/env python3
"""
Test the fixed API with Sierra input file
"""
import requests
import json

def test_api():
    print("=== TESTING FIXED API ===")
    
    # API base URL 
    base_url = "https://9000-idrlbzy4bg2q93rmh2rr0-6532622b.e2b.dev"
    
    # Test health endpoint
    print("1. Testing health endpoint...")
    health_response = requests.get(f"{base_url}/api/health")
    print(f"Health Status: {health_response.status_code}")
    if health_response.status_code == 200:
        print(f"Health Data: {health_response.json()}")
    
    # Test file upload and conversion
    print("\n2. Testing file upload and conversion...")
    
    # Upload Sierra file
    files = {
        'file': open('/home/user/webapp/sierra_input_new.xlsx', 'rb')
    }
    
    upload_response = requests.post(f"{base_url}/api/convert", files=files)
    print(f"Upload Status: {upload_response.status_code}")
    
    if upload_response.status_code == 200:
        result = upload_response.json()
        print(f"Conversion Result: {json.dumps(result, indent=2)}")
        
        # Download the converted file if available
        if result.get('success') and 'download_url' in result:
            print(f"\n3. Downloading converted file...")
            download_url = result['download_url']
            if not download_url.startswith('http'):
                download_url = base_url + download_url
                
            download_response = requests.get(download_url)
            if download_response.status_code == 200:
                with open('/home/user/webapp/api_test_output.xlsx', 'wb') as f:
                    f.write(download_response.content)
                print(f"✅ Downloaded API output to: api_test_output.xlsx")
            else:
                print(f"❌ Download failed: {download_response.status_code}")
        
    else:
        print(f"❌ Conversion failed: {upload_response.text}")
    
    files['file'].close()

if __name__ == "__main__":
    test_api()