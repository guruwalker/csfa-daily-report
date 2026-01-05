"""
API Client for CSFA Report Automation
Handles all API interactions with proper error handling and validation.
"""

import requests
import os
import logging
from dotenv import load_dotenv
from typing import Dict, Any, Optional

# Load .env file
load_dotenv()

logger = logging.getLogger(__name__)

BASE_URL = "https://tintasberger.solutechlabs.com"


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def clean_token(token: str) -> str:
    """
    Clean and validate token string.

    Args:
        token: Raw token string from environment

    Returns:
        Cleaned token string

    Raises:
        ValueError: If token is invalid
    """
    if not token:
        raise ValueError("Token is empty or None")

    # Convert to string if it's bytes
    if isinstance(token, bytes):
        token = token.decode('utf-8')

    # Strip whitespace and quotes
    token = str(token).strip().strip('"').strip("'")

    # Remove any control characters
    token = ''.join(char for char in token if ord(char) >= 32)

    # Check if token was masked (common in CI/CD)
    if token == '***' or token.startswith('***'):
        raise ValueError(
            "Token appears to be masked. "
            "Ensure secrets are properly configured in your CI/CD environment."
        )

    # Validate token format (basic check)
    if len(token) < 10:
        raise ValueError(f"Token appears invalid (too short: {len(token)} characters)")

    return token


def get_validated_token() -> str:
    """
    Get and validate the access token from environment.

    Returns:
        Validated access token

    Raises:
        ValueError: If token is missing or invalid
    """
    token = os.getenv("ACCESS_TOKEN")

    if not token:
        raise ValueError(
            "ACCESS_TOKEN not found in environment variables. "
            "Please ensure your .env file is configured correctly."
        )

    try:
        return clean_token(token)
    except ValueError as e:
        logger.error(f"Token validation failed: {e}")
        raise


# ============================================================================
# API FUNCTIONS
# ============================================================================

def get_orders(access_token: str, query_string: str) -> Dict[str, Any]:
    """
    Fetch orders from the API.

    Args:
        access_token: API bearer token
        query_string: Query string with parameters

    Returns:
        JSON response containing orders data

    Raises:
        requests.RequestException: If API request fails
    """
    try:
        # Clean and validate token
        clean_access_token = clean_token(access_token)

        url = f"{BASE_URL}/api/v1/get-v2-orders{query_string}"
        headers = {
            "Authorization": f"Bearer {clean_access_token}",
            "Accept": "application/json",
        }

        logger.debug(f"Fetching orders from: {url}")

        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        return response.json()

    except ValueError as e:
        logger.error(f"Token validation error: {e}")
        raise
    except requests.RequestException as e:
        logger.error(f"API request failed: {e}")
        raise


def get_timesheet(headers: dict, cookies: dict, params: dict) -> Dict[str, Any]:
    """
    Fetch timesheet data from the API.

    Args:
        headers: Request headers
        cookies: Request cookies
        params: Query parameters

    Returns:
        JSON response containing timesheet data

    Raises:
        requests.RequestException: If API request fails
    """
    try:
        # Clean headers that might contain tokens
        cleaned_headers = {}
        for key, value in headers.items():
            if isinstance(value, (str, bytes)):
                try:
                    cleaned_value = clean_token(value) if 'token' in key.lower() else str(value).strip()
                except ValueError:
                    # If cleaning fails, use original but ensure it's a string
                    cleaned_value = str(value).strip()
                cleaned_headers[key] = cleaned_value
            else:
                cleaned_headers[key] = value

        url = f"{BASE_URL}/timesheet-list"

        logger.debug(f"Fetching timesheet from: {url}")

        response = requests.get(
            url,
            headers=cleaned_headers,
            cookies=cookies,
            params=params,
            timeout=30
        )
        response.raise_for_status()

        return response.json()

    except ValueError as e:
        logger.error(f"Header validation error: {e}")
        raise
    except requests.RequestException as e:
        logger.error(f"API request failed: {e}")
        raise


def get_order_details(access_token: str, order_number: int | str) -> Dict[str, Any]:
    """
    Fetch details for a single order by order number.

    Args:
        access_token: API bearer token
        order_number: Order number to fetch details for

    Returns:
        JSON response containing order details

    Raises:
        requests.RequestException: If API request fails
    """
    try:
        # Clean and validate token
        clean_access_token = clean_token(access_token)

        url = f"{BASE_URL}/api/v1/get-v2-order-details/{order_number}"
        headers = {
            "Authorization": f"Bearer {clean_access_token}",
            "Accept": "application/json",
        }

        logger.debug(f"Fetching order details for order: {order_number}")

        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        return response.json()

    except ValueError as e:
        logger.error(f"Token validation error: {e}")
        raise
    except requests.RequestException as e:
        logger.error(f"API request failed for order {order_number}: {e}")
        raise


# ============================================================================
# VALIDATION ON MODULE LOAD
# ============================================================================

# Validate token when module is imported (fail fast)
try:
    _validated_token = get_validated_token()
    logger.info("✅ Access token validated successfully")
except ValueError as e:
    logger.warning(f"⚠️ Token validation issue on module load: {e}")
    logger.warning("This may be expected in CI/CD environments with masked secrets")
