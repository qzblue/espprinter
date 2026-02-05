"""
LDAP Service for Active Directory User Lookup

This module provides functionality to query Microsoft Active Directory
for user display names. It includes caching to minimize LDAP queries.
"""

import os
import logging
from typing import Optional
from functools import lru_cache

try:
    from ldap3 import Server, Connection, ALL, SUBTREE
    from ldap3.core.exceptions import LDAPException
    LDAP_AVAILABLE = True
except ImportError:
    LDAP_AVAILABLE = False

# Configure logging
logger = logging.getLogger(__name__)

# LDAP Configuration from environment variables
LDAP_URL = os.getenv("LDAP_URL", "ldap://10.32.64.200:389")
LDAP_ADMIN_DN = os.getenv("LDAP_ADMIN_DN", "espsupmu")
LDAP_ADMIN_PASSWORD = os.getenv("LDAP_ADMIN_PASSWORD", "$@intP@u1$ch001M@C@U@2018")
LDAP_BASE_DN = os.getenv("LDAP_BASE_DN", "DC=esptals,DC=esp,DC=edu,DC=mo")
LDAP_USER_SEARCH_FILTER = os.getenv("LDAP_USER_SEARCH_FILTER", "(samAccountName={0})")


def _create_ldap_connection() -> Optional[Connection]:
    """
    Create and bind an LDAP connection to Active Directory.
    
    Returns:
        Connection object if successful, None otherwise
    """
    if not LDAP_AVAILABLE:
        logger.warning("ldap3 library not available. Install with: pip install ldap3")
        return None
    
    try:
        server = Server(LDAP_URL, get_info=ALL)
        conn = Connection(
            server,
            user=LDAP_ADMIN_DN,
            password=LDAP_ADMIN_PASSWORD,
            auto_bind=True
        )
        return conn
    except LDAPException as e:
        logger.error(f"Failed to connect to LDAP server: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error connecting to LDAP: {e}")
        return None


@lru_cache(maxsize=1000)
def get_user_display_name(username: str) -> str:
    """
    Fetch user's display name from Active Directory.
    
    Results are cached in memory for performance. Cache is cleared
    when the application restarts.
    
    Args:
        username: The samAccountName (username) to look up
        
    Returns:
        Display name if found, otherwise returns the original username
        
    Examples:
        >>> get_user_display_name("psi.teacher01")
        "João Silva"
        >>> get_user_display_name("unknown.user")
        "unknown.user"
    """
    if not username or not username.strip():
        return username or "未知"
    
    username = username.strip()
    
    # Quick return if LDAP is not available
    if not LDAP_AVAILABLE:
        return username
    
    conn = None
    try:
        conn = _create_ldap_connection()
        if not conn:
            return username
        
        # Search for the user
        search_filter = LDAP_USER_SEARCH_FILTER.format(username)
        conn.search(
            search_base=LDAP_BASE_DN,
            search_filter=search_filter,
            search_scope=SUBTREE,
            attributes=['displayName', 'cn', 'name']
        )
        
        if conn.entries:
            entry = conn.entries[0]
            # Try different attributes in order of preference
            if hasattr(entry, 'displayName') and entry.displayName:
                return str(entry.displayName.value)
            elif hasattr(entry, 'cn') and entry.cn:
                return str(entry.cn.value)
            elif hasattr(entry, 'name') and entry.name:
                return str(entry.name.value)
        
        # No match found, return original username
        logger.debug(f"No LDAP entry found for username: {username}")
        return username
        
    except LDAPException as e:
        logger.error(f"LDAP query failed for user '{username}': {e}")
        return username
    except Exception as e:
        logger.error(f"Unexpected error querying LDAP for '{username}': {e}")
        return username
    finally:
        if conn:
            try:
                conn.unbind()
            except:
                pass


def format_user_display(username: str, show_username: bool = True) -> str:
    """
    Format user display string with display name and optional username.
    
    Args:
        username: The username to format
        show_username: If True, append username in parentheses
        
    Returns:
        Formatted display string
        
    Examples:
        >>> format_user_display("psi.teacher01", True)
        "João Silva (psi.teacher01)"
        >>> format_user_display("psi.teacher01", False)
        "João Silva"
    """
    display_name = get_user_display_name(username)
    
    # If display name is same as username, just return it
    if display_name == username:
        return username
    
    # Return with or without username suffix
    if show_username:
        return f"{display_name} ({username})"
    else:
        return display_name




@lru_cache(maxsize=500)
def search_usernames_by_display_name(display_name_query: str) -> tuple:
    """
    Search LDAP for usernames matching a display name query.
    
    This enables searching by Chinese names, English names, or partial matches.
    Results are cached for performance.
    
    Args:
        display_name_query: The display name or partial name to search for
        
    Returns:
        Tuple of matching usernames (samAccountName values)
        
    Examples:
        >>> search_usernames_by_display_name("張穎儀")
        ('cwy0310',)
        >>> search_usernames_by_display_name("Wing-Yee")
        ('cwy0310',)
    """
    if not display_name_query or not display_name_query.strip():
        return ()
    
    display_name_query = display_name_query.strip()
    
    # Quick return if LDAP is not available
    if not LDAP_AVAILABLE:
        return ()
    
    conn = None
    try:
        conn = _create_ldap_connection()
        if not conn:
            return ()
        
        # Build LDAP search filter for display name, cn, or name fields
        # Use wildcard search for partial matches
        search_filter = f"(|(displayName=*{display_name_query}*)(cn=*{display_name_query}*)(name=*{display_name_query}*))"
        
        conn.search(
            search_base=LDAP_BASE_DN,
            search_filter=search_filter,
            search_scope=SUBTREE,
            attributes=['samAccountName'],
            size_limit=50  # Limit results to prevent too many matches
        )
        
        if conn.entries:
            usernames = []
            for entry in conn.entries:
                if hasattr(entry, 'samAccountName') and entry.samAccountName:
                    username = str(entry.samAccountName.value)
                    usernames.append(username)
            
            logger.debug(f"Found {len(usernames)} users matching '{display_name_query}'")
            return tuple(usernames)  # Return tuple for hashability (caching)
        
        # No match found
        logger.debug(f"No LDAP entries found for display name query: {display_name_query}")
        return ()
        
    except LDAPException as e:
        logger.error(f"LDAP search failed for display name '{display_name_query}': {e}")
        return ()
    except Exception as e:
        logger.error(f"Unexpected error searching LDAP for '{display_name_query}': {e}")
        return ()
    finally:
        if conn:
            try:
                conn.unbind()
            except:
                pass


def clear_cache():
    """Clear the LDAP lookup cache. Useful for testing or if AD data changes."""
    get_user_display_name.cache_clear()
    search_usernames_by_display_name.cache_clear()
    logger.info("LDAP cache cleared")
