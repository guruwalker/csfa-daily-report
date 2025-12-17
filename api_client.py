import requests
import os
from dotenv import load_dotenv

# Load .env file
load_dotenv()


access_token = os.getenv("ACCESS_TOKEN")

# print(f"API Client Access Token: {access_token}")

BASE_URL = "https://tintasberger.solutechlabs.com"

# --- Existing functions ---
def get_orders(access_token, query_string):
    url = f"{BASE_URL}/api/v1/get-v2-orders{query_string}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def get_timesheet(headers: dict, cookies: dict, params: dict):
    url = f"{BASE_URL}/timesheet-list"
    response = requests.get(url, headers=headers, cookies=cookies, params=params)
    response.raise_for_status()
    return response.json()


# --- New function: get single order details ---
def get_order_details(access_token, order_number):
    """
    Fetch details for a single order by order number.

    Args:
        access_token (str): API bearer token
        order_number (int or str): Order number to fetch details for

    Returns:
        dict: JSON response containing order details
    """
    url = f"https://tintasberger.solutechlabs.com/api/v1/get-v2-order-details/{order_number}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()
