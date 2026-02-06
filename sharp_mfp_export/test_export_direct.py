import sys
sys.path.insert(0, '.')

from webapp import app

# Enable debug mode temporarily
app.config['TESTING'] = True

with app.test_client() as client:
    print("Testing /export/jobs endpoint...")
    try:
        response = client.get('/export/jobs')
        print(f"Status: {response.status_code}")
        
        if response.status_code == 200:
            print(f"Success! File size: {len(response.data)} bytes")
        else:
            print("Error!")
            print(response.data.decode('utf-8'))
            
    except Exception as e:
        print(f"Exception occurred: {e}")
        import traceback
        traceback.print_exc()
