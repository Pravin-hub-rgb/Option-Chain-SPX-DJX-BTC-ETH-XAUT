import os
import tempfile
import unittest

from product_identity import (
    PRODUCT_ID,
    PRODUCT_MARKER,
    identify_product_directory,
    same_product_directory,
)


class ProductIdentityTests(unittest.TestCase):
    def test_marker_identifies_current_product(self):
        with tempfile.TemporaryDirectory() as directory:
            with open(os.path.join(directory, PRODUCT_MARKER), "w", encoding="ascii") as marker:
                marker.write(PRODUCT_ID)
            self.assertEqual(identify_product_directory(directory), "us")

    def test_legacy_workbook_set_identifies_us_without_marker(self):
        with tempfile.TemporaryDirectory() as directory:
            for filename in (
                "spx_option_chain.xlsx",
                "djx_option_chain.xlsx",
                "us_stock_option_chain.xlsx",
            ):
                with open(os.path.join(directory, filename), "wb"):
                    pass
            self.assertEqual(identify_product_directory(directory), "us")

    def test_legacy_delta_workbook_set_is_not_misidentified_as_us(self):
        with tempfile.TemporaryDirectory() as directory:
            for filename in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
                with open(os.path.join(directory, filename), "wb"):
                    pass
            self.assertEqual(identify_product_directory(directory), "delta")

    def test_unknown_or_partial_directory_is_left_unidentified(self):
        with tempfile.TemporaryDirectory() as directory:
            with open(os.path.join(directory, "spx_option_chain.xlsx"), "wb"):
                pass
            self.assertIsNone(identify_product_directory(directory))

    def test_invalid_marker_does_not_fall_back_to_workbooks(self):
        with tempfile.TemporaryDirectory() as directory:
            for filename in (
                "spx_option_chain.xlsx",
                "djx_option_chain.xlsx",
                "us_stock_option_chain.xlsx",
            ):
                with open(os.path.join(directory, filename), "wb"):
                    pass
            with open(os.path.join(directory, PRODUCT_MARKER), "w", encoding="ascii") as marker:
                marker.write("unknown")
            self.assertIsNone(identify_product_directory(directory))

    def test_cross_folder_us_copy_replaces_legacy_us_only(self):
        with tempfile.TemporaryDirectory() as first, tempfile.TemporaryDirectory() as legacy:
            for directory in (first, legacy):
                for filename in (
                    "spx_option_chain.xlsx",
                    "djx_option_chain.xlsx",
                    "us_stock_option_chain.xlsx",
                ):
                    with open(os.path.join(directory, filename), "wb"):
                        pass
            self.assertTrue(same_product_directory(first, legacy))

    def test_us_copy_does_not_replace_delta_copy(self):
        with tempfile.TemporaryDirectory() as us, tempfile.TemporaryDirectory() as delta:
            for filename in (
                "spx_option_chain.xlsx",
                "djx_option_chain.xlsx",
                "us_stock_option_chain.xlsx",
            ):
                with open(os.path.join(us, filename), "wb"):
                    pass
            for filename in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
                with open(os.path.join(delta, filename), "wb"):
                    pass
            self.assertFalse(same_product_directory(us, delta))

    def test_unknown_directory_is_not_a_same_product_match(self):
        with tempfile.TemporaryDirectory() as us, tempfile.TemporaryDirectory() as unknown:
            for filename in (
                "spx_option_chain.xlsx",
                "djx_option_chain.xlsx",
                "us_stock_option_chain.xlsx",
            ):
                with open(os.path.join(us, filename), "wb"):
                    pass
            self.assertFalse(same_product_directory(us, unknown))


if __name__ == "__main__":
    unittest.main()
