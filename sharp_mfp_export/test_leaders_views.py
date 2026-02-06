import requests
from bs4 import BeautifulSoup

# Test the three view modes
modes = ["all_printers", "single_printer", "aggregated"]

for mode in modes:
    url = f"http://127.0.0.1:5000/leaders?view_mode={mode}"
    print(f"\n{'='*60}")
    print(f"Testing {mode} mode")
    print(f"URL: {url}")
    print(f"{'='*60}")
    
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Check view mode selector exists
            view_select = soup.find('select', id='view-mode-select')
            if view_select:
                selected = view_select.find('option', selected=True)
                print(f"✓ View mode selector found, current: {selected.get('value') if selected else 'none'}")
            
            # Check printer selector visibility
            printer_label = soup.find('label', id='printer-selector')
            if printer_label:
                style = printer_label.get('style', '')
                visible = 'none' not in style.lower() or mode == 'single_printer'
                print(f"✓ Printer selector visibility: {'shown' if 'none' not in style.lower() else 'hidden'} (expected: {'shown' if mode == 'single_printer' else 'hidden'})")
            
            # Count printer cards
            cards = soup.find_all('section', class_='card')
            header_cards = [c for c in cards if 'printer' in str(c.find('h2'))]
            print(f"✓ Found {len(header_cards)} printer card(s)")
            
            # Check for aggregated section
            agg = soup.find('h2', string=lambda t: t and '跨列印機' in t)
            if agg:
                print(f"✓ Aggregated section found")
            
            # Check table headers
            tables = soup.find_all('table')
            for i, table in enumerate(tables[:2]):  # Check first 2 tables
                headers = [th.get_text(strip=True) for th in table.find_all('th')]
                print(f"  Table {i+1} headers: {headers}")
            
            print(f"✓ Status: Success")
        else:
            print(f"✗ HTTP {response.status_code}")
    except Exception as e:
        print(f"✗ Error: {e}")

print(f"\n{'='*60}")
print("Test Complete")
print(f"{'='*60}")
