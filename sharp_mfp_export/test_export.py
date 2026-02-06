import requests

url = "http://127.0.0.1:5000/export/jobs"
print(f"Testing: {url}")

try:
    response = requests.get(url, timeout=30)
    print(f"Status Code: {response.status_code}")
    print(f"Headers: {response.headers}")
    
    if response.status_code == 200:
        print(f"Success! File size: {len(response.content)} bytes")
        # Save to file for inspection
        with open("test_export.xlsx", "wb") as f:
            f.write(response.content)
        print("Saved to test_export.xlsx")
    else:
        print(f"Error Response:")
        print(response.text[:1000])
        
except Exception as e:
    print(f"Exception: {e}")
    import traceback
    traceback.print_exc()
