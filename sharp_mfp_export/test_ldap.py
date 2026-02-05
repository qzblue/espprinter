import os
import sys
import logging

# Configure logging to stdout
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

print("--- LDAP Diagnostic Tool ---")

try:
    import ldap_service
    print(f"Successfully imported ldap_service")
except ImportError as e:
    print(f"CRITICAL: Failed to import ldap_service: {e}")
    sys.exit(1)

print(f"LDAP_AVAILABLE: {ldap_service.LDAP_AVAILABLE}")

if not ldap_service.LDAP_AVAILABLE:
    print("❌ LDAP library (ldap3) is NOT installed or failed to load.")
    print("Please verify 'ldap3' is in requirements.txt and installed in the Docker image.")
    sys.exit(1)
    
print(f"LDAP Config:")
print(f"  URL: {ldap_service.LDAP_URL}")
print(f"  Base DN: {ldap_service.LDAP_BASE_DN}")
print(f"  Admin DN: {ldap_service.LDAP_ADMIN_DN}")
print(f"  Password: {'*' * len(ldap_service.LDAP_ADMIN_PASSWORD) if ldap_service.LDAP_ADMIN_PASSWORD else 'None'}")

print("\n--- Testing Connection ---")
conn = ldap_service._create_ldap_connection()

if conn:
    print("✅ Connection Successful!")
    print(f"  Server Info: {conn.server}")
    print(f"  User: {conn.user}")
    
    print("\n--- Testing Lookup ---")
    # Try to find a known user, or just query for * something small
    test_user = "espsupmu" # Self lookup usually works if admin
    print(f"Looking up user '{test_user}'...")
    display_name = ldap_service.get_user_display_name(test_user)
    print(f"Result: '{display_name}'")
    
    if display_name == test_user:
         print("⚠️ Lookup returned username. This implies user not found or no display name attribute.")
    else:
         print("✅ Lookup successful! Display name found.")
         
    conn.unbind()
else:
    print("❌ Connection Failed.")
    print("Please check: ")
    print("1. Network connectivity to the LDAP server IP (firewall/routing).")
    print("2. Validity of Admin DN and Password.")

print("\n--- Diagnostic Complete ---")
