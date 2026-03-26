import requests
import pandas as pd

# ---------------- CONFIG ----------------
ACCESS_TOKEN="eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiYzYxMmY5NzJhYTgyZDhlNjY2ZTc0MjhiMTgyNWIxZThkNjdhNmE3YjU2YzkwYTQxOTBhMzE5N2MzZTZlMzQ0OWZjODAyZjUzMzE3MWViZDciLCJpYXQiOjE3NjQ2NTQ4NzAuMzcxNDU5LCJuYmYiOjE3NjQ2NTQ4NzAuMzcxNDYzLCJleHAiOjE3OTYxOTA4NzAuMzYwMywic3ViIjoiNTciLCJzY29wZXMiOltdfQ.AQ-5l2TZ6h-4-ijgXvfbRCrqDwR8YIrSq0bOP3WpGHVXiwRSy5ffY48n9YmJeLI50jpdVgkotfXQuap7p9P9kI5-SimDZ638GUv0BePKbRYzVQJcMwHeyrb4OJaP1FDCp848w4bdGOuvuQAfoCs0JyoQ6ngxC1X_qmHzjdHStd4uh7rF4elljstEF6LWe4BXb2vVY87rbof_mQCbdxg47D1s5ToQNIHPcYtCq64nkk-5Rol0fJoE1xffLAhDdUuyv9FGStuOPY9D-VY-sgOoeTGMF9eLgVBZXWZKVYptR0EOALpyn5MqOxGsECw7JU004AlsBP-Ur6JEpIG1H6yk_c-plJgQU3DXReJp-62w8P7QJYmjd3-h1H850OF_N_ODtfz5sjYBs7JSG1DyzXEUg5_xXV-7hnHtSTF440FBUlvhyY_Ki4N2a45MtET7EHA2-rOACMxf1lAPcjVufQmqDiTmbdsAXq-j2_krn57whC5UXJOMlsNXRY9RdaI19-P9QdO5eOvQRb0Tj7GIUDGoktqj19_I5DBloh2dFS1UBxhrFHZM5bVUE2RN5C1zZuNQu1nMsO5ousflTtYadQMVIVN4EOmd5cdgfJt5Fgan3bJoquEzcZSDTJuLGiPyj-Xe5_gBRGCBk-XmsiDaJg4MIplVVH-9jSaW00tpW7AMMgw"
START_ORDER = 3395
END_ORDER = 3455
OUTPUT_FILE = "all_orders.xlsx"

BASE_URL = "https://tintasberger.solutechlabs.com"
# --------------------------------------


def get_order_sku_details(access_token, order_number):
    url = f"{BASE_URL}/api/v1/get-v2-order-sku-details/{order_number}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()


def main():
    rows = []
    failed_orders = []

    for order_number in range(START_ORDER, END_ORDER + 1):
        try:
            print(f"Fetching order {order_number}...")

            response = get_order_sku_details(ACCESS_TOKEN, order_number)

            data = response.get("data", {})

            # ✅ CUSTOMER NAME (CORRECT SOURCE)
            customer_name = (
                data.get("customer", {}).get("shop_name")
                if isinstance(data.get("customer"), dict)
                else None
            )

            order_details = data.get("order_details", [])

            if not isinstance(order_details, list) or not order_details:
                print(f"⚠️ No order details for order {order_number}")
                continue

            for item in order_details:
                rows.append({
                    "order_number": order_number,
                    "customer_name": customer_name,
                    "order_id": item.get("sales_order_detail_id"),
                    "product_id": item.get("product_id"),
                    "product_description": item.get("product_desc"),
                    "item_description": item.get("product_ref"),
                    "sold_qty": item.get("quantity"),
                    "uom": item.get("package_name"),
                })

        except Exception as e:
            print(f"❌ Failed order {order_number}: {e}")
            failed_orders.append(order_number)

    df = pd.DataFrame(rows)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Order Lines")

        if failed_orders:
            pd.DataFrame(
                {"failed_order_number": failed_orders}
            ).to_excel(writer, index=False, sheet_name="Failed Orders")

    print("\n✅ DONE")
    print(f"Excel created: {OUTPUT_FILE}")
    print(f"Rows written: {len(df)}")

    if failed_orders:
        print("⚠️ Failed orders:", failed_orders)


if __name__ == "__main__":
    main()
