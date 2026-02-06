import requests
import sys

BASE_URL = "http://127.0.0.1:5000"

def test_view(name, params, checks):
    print(f"Testing {name}...")
    try:
        resp = requests.get(f"{BASE_URL}/leaders", params=params)
        resp.raise_for_status()
        html = resp.text
        
        for check_name, check_str, should_exist in checks:
            if should_exist:
                if check_str in html:
                    print(f"  [PASS] {check_name}: Found '{check_str}'")
                else:
                    print(f"  [FAIL] {check_name}: NOT Found '{check_str}'")
                    # print(html) # Debug
                    return False
            else:
                if check_str not in html:
                    print(f"  [PASS] {check_name}: Correctly NOT Found '{check_str}'")
                else:
                    print(f"  [FAIL] {check_name}: Found unexpectedly '{check_str}'")
                    return False
        return True
    except Exception as e:
        print(f"  [ERROR] {e}")
        return False

def test_export():
    print("Testing Export (All Data)...")
    try:
        resp = requests.get(f"{BASE_URL}/export/leaders", params={"export_range": "all_data"})
        if resp.status_code == 200:
             print(f"  [PASS] Status code 200")
        else:
             print(f"  [FAIL] Status code {resp.status_code}")
             return False
             
        ct = resp.headers.get("Content-Type", "")
        if "spreadsheetml" in ct or "excel" in ct:
            print(f"  [PASS] Content-Type seems correct: {ct}")
        else:
            print(f"  [WARN] Content-Type: {ct}")
            
        cd = resp.headers.get("Content-Disposition", "")
        if "leaders_all_data.xlsx" in cd:
             print(f"  [PASS] Filename correct: {cd}")
        else:
             print(f"  [FAIL] Filename not in {cd}")
             return False
        return True
    except Exception as e:
        print(f"  [ERROR] {e}")
        return False

def main():
    # 1. All Printers
    # Expect "列印機" column
    success = test_view("All Printers", {"view_mode": "all_printers", "printer": "http://10.64.48.120"}, [
        ("Printer Column", "<th>列印機</th>", True),
        ("Table Row", "<tr>", True),
        ("Unified Header", "<h2>排行榜", True)
    ])
    
    # 2. Single Printer
    # Expect NO "列印機" column
    success &= test_view("Single Printer", {"view_mode": "single_printer", "printer": "http://10.64.48.120"}, [
        ("Printer Column", "<th>列印機</th>", False),
        ("Table Row", "<tr>", True)
    ])
    
    # 3. Aggregated
    # Expect NO "列印機" column
    success &= test_view("Aggregated", {"view_mode": "aggregated"}, [
        ("Printer Column", "<th>列印機</th>", False),
        ("Table Row", "<tr>", True)
    ])
    
    # 4. Export
    success &= test_export()
    
    if success:
        print("\nALL TESTS PASSED")
    else:
        print("\nSOME TESTS FAILED")
        sys.exit(1)

if __name__ == "__main__":
    main()
