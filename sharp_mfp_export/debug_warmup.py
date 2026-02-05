
import os
import sys
from sharp_mfp_export import warmup_webapp

# Set env var for testing
os.environ["WEBAPP_URL"] = "http://localhost:5000"

print("Testing warmup_webapp()...")
try:
    warmup_webapp()
    print("Test Complete.")
except Exception as e:
    print(f"Test Failed: {e}")
