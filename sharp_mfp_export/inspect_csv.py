
import glob
import csv
import sys

files = glob.glob('exports/joblog/*.csv')
if not files:
    print("No CSV files found.")
    sys.exit(0)

target = files[0]
print(f"Reading {target}...")

with open(target, 'r', encoding='big5', errors='replace') as f:
    reader = csv.DictReader(f)
    print("Headers:", reader.fieldnames)
    
    count = 0
    for row in reader:
        # Check for print mode and file name
        mode = row.get("工作模式") or row.get("Job Mode") or row.get("Mode")
        fname = row.get("檔案名稱")
        ftype = row.get("檔案類型")
        
        if fname or ftype:
             print(f"Row {count}: Mode={mode}, Name={fname}, Type={ftype}")
             count += 1
             if count >= 10:
                 break
