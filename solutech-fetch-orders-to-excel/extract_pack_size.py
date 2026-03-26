import pandas as pd
import re


def extract_packsize(product_description: str):
    """
    Extract pack size from product_description and return 3-digit code.
    Examples:
        1L     -> 001
        5L     -> 005
        20L    -> 020
        0.5L   -> 500
        0.75L  -> 750
        30KG   -> 030
        25KG   -> 025
        5KG    -> 005
    """

    if not isinstance(product_description, str):
        return None

    match = re.search(r"(\d+(?:\.\d+)?)\s*(L|KG)", product_description.upper())
    if not match:
        return None

    value = float(match.group(1))
    unit = match.group(2)

    if unit == "L":
        if value < 1:
            code = int(value * 1000)   # 0.5L → 500
        else:
            code = int(value)          # 5L → 5
    else:  # KG
        code = int(value)

    return f"{code:03d}"


def build_packsize_file(input_file, output_file):
    df = pd.read_excel(input_file)

    packsize_values = []

    for _, row in df.iterrows():
        item_desc = row.get("item_description")
        product_desc = row.get("product_description")

        size_code = extract_packsize(product_desc)

        if size_code and isinstance(item_desc, str):
            packsize = f"{item_desc}-{size_code}"
        else:
            packsize = None

        packsize_values.append(packsize)

    df["packsize"] = packsize_values

    df.to_excel(output_file, index=False)

    print("✅ Packsize extraction complete")
    print(f"Rows processed: {len(df)}")
    print(f"Output file: {output_file}")


# ---------------- RUN ----------------

build_packsize_file(
    input_file="all_orders.xlsx",
    output_file="all_orders_with_packsize.xlsx"
)
