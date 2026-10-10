"""Identify Option Chain package directories without relying on their install path."""

import os


PRODUCT_MARKER = "optionchain.product"
PRODUCT_ID = "us"
PRODUCT_WORKBOOKS = {
    "delta": ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"),
    "us": (
        "spx_option_chain.xlsx",
        "djx_option_chain.xlsx",
        "us_stock_option_chain.xlsx",
    ),
}


def identify_product_directory(directory):
    """Return a known product ID, or None when the directory is ambiguous."""
    marker_path = os.path.join(directory, PRODUCT_MARKER)
    if os.path.exists(marker_path):
        try:
            with open(marker_path, encoding="ascii") as marker:
                product_id = marker.read().strip().lower()
        except OSError:
            return None
        return product_id if product_id in PRODUCT_WORKBOOKS else None

    matches = [
        product_id
        for product_id, workbooks in PRODUCT_WORKBOOKS.items()
        if all(os.path.isfile(os.path.join(directory, workbook)) for workbook in workbooks)
    ]
    return matches[0] if len(matches) == 1 else None


def same_product_directory(first_directory, second_directory):
    first_product = identify_product_directory(first_directory)
    return (
        first_product is not None
        and first_product == identify_product_directory(second_directory)
    )
