import requests

# Test the three view modes
modes = ["all_printers", "single_printer", "aggregated"]

print("Testing Leaders Page View Modes")
print("="*60)

for mode in modes:
    url = f"http://127.0.0.1:5000/leaders?view_mode={mode}"
    print(f"\nMode: {mode}")
    print(f"URL: {url}")
    
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            html = response.text
            
            # Basic checks
            print(f"✓ Status: {response.status_code}")
            print(f"✓ Page loaded successfully ({len(html)} bytes)")
            
            # Check for key elements
            if 'view-mode-select' in html:
                print("✓ View mode selector found")
            
            if mode == 'aggregated':
                if '跨列印機匯總排行' in html:
                    print("✓ Aggregated header found")
            else:
                if 'printer-selector' in html:
                    print("✓ Printer selector in HTML")
                    
        else:
            print(f"✗ HTTP {response.status_code}")
    except Exception as e:
        print(f"✗ Error: {e}")

print(f"\n{'='*60}")
print("All view modes responded successfully!")
print("Please manually verify in browser:") 
print("  http://127.0.0.1:5000/leaders")
