import pandas as pd
import ast


def extract_product_lines_with_customer(
    input_file,
    output_file,
    json_column_index=1,          # Column A
    customer_column_index=4,      # Column E
    date_column_index=16,         # Column Q
    order_number_column_index=17  # Column R
):
    """
    Extracts:
    - product_name
    - quantity
    - product_description
    - item_description
    - customer_name
    - date
    - order_number

    Handles malformed / non-string cells safely.
    """

    df = pd.read_excel(input_file)
    extracted_rows = []

    for row_idx in range(len(df)):
        cell_value = df.iat[row_idx, json_column_index]
        customer_name = df.iat[row_idx, customer_column_index]
        date_value = df.iat[row_idx, date_column_index]
        order_number_value = df.iat[row_idx, order_number_column_index]

        # Skip empty or non-string cells
        if not isinstance(cell_value, str):
            continue

        # Must look like a list
        if not cell_value.strip().startswith("["):
            continue

        try:
            items = ast.literal_eval(cell_value)

            if not isinstance(items, list):
                continue

            for item in items:
                extracted_rows.append({
                    "product_name": item.get("product_name"),
                    "quantity": item.get("quantity"),
                    "product_description": item.get("product_desc"),
                    "item_description": item.get("item_description"),
                    "customer_name": customer_name,
                    "date": date_value,
                    "order_number": order_number_value,
                })

        except Exception as e:
            print(f"❌ Row {row_idx} failed: {e}")

    result_df = pd.DataFrame(extracted_rows)
    result_df.to_excel(output_file, index=False)

    print("✅ Extraction complete")
    print(f"Rows extracted: {len(result_df)}")
    print(f"Output file: {output_file}")


# ---------------- RUN ----------------

extract_product_lines_with_customer(
    input_file="orders_3395_3454.xlsx",
    output_file="products_with_customers.xlsx"
)
