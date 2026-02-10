"""
Configuration file for CSFA Report
Contains list of all salespeople and filtering rules.
"""

# ============================================================================
# SALESPEOPLE CONFIGURATION
# ============================================================================

# Complete list of all salespeople who should be using the app
ALL_SALESPEOPLE = [
    "ANDRE MARQUEZA",
    "IMRAN AHMED",
    "INÁCIO RODRIGUES",
    "KUMAR CHAMPAKLAL",
    # "LAURA MARCIA",
    "RICARDO DINIS LANGA",
    # "ITO BEDITO",
    # "FRANCISCO DO ROSARIO TOMAS",
    # "SAIDATA ZALIA JAUHAR SAIDE",
    # "HENRIQUE BERTUR MARCO",
    "RICARDO MACUACUA"
]

# Accounts to exclude from all reports (test accounts, etc.)
EXCLUDED_ACCOUNTS = [
    "TEST ACCOUNT",
    "Test Account",
    "test account",
]


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def normalize_name(name: str) -> str:
    """
    Normalize salesperson name for comparison.

    Args:
        name: Original name

    Returns:
        Normalized name (trimmed, title case)
    """
    if not name:
        return ""
    return name.strip()


def is_excluded_account(name: str) -> bool:
    """
    Check if an account should be excluded from reports.

    Args:
        name: Account/salesperson name

    Returns:
        True if account should be excluded
    """
    if not name:
        return False

    normalized = name.strip().lower()

    for excluded in EXCLUDED_ACCOUNTS:
        if excluded.lower() in normalized or normalized in excluded.lower():
            return True

    return False


def get_all_salespeople() -> list:
    """
    Get list of all salespeople.

    Returns:
        List of all salesperson names
    """
    return ALL_SALESPEOPLE.copy()


def filter_excluded_accounts(data_list: list, name_key: str = "name") -> list:
    """
    Filter out excluded accounts from a list of dictionaries.

    Args:
        data_list: List of dictionaries containing data
        name_key: Key to use for name lookup

    Returns:
        Filtered list without excluded accounts
    """
    return [
        item for item in data_list
        if not is_excluded_account(item.get(name_key, ""))
    ]
