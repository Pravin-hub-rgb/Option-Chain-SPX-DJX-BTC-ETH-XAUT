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
            marker_path = os.path.join(directory, PRODUCT_MARKER)
            with open(marker_path, "w", encoding="ascii") as marker:
                marker.write(PRODUCT_ID)
            self.assertEqual(identify_product_directory(directory), "delta")

    def test_legacy_workbook_set_identifies_delta_without_marker(self):
        with tempfile.TemporaryDirectory() as directory:
            for filename in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
                with open(os.path.join(directory, filename), "wb"):
                    pass
            self.assertEqual(identify_product_directory(directory), "delta")

    def test_legacy_us_workbook_set_is_not_misidentified_as_delta(self):
        with tempfile.TemporaryDirectory() as directory:
            for filename in (
                "spx_option_chain.xlsx",
                "djx_option_chain.xlsx",
                "us_stock_option_chain.xlsx",
            ):
                with open(os.path.join(directory, filename), "wb"):
                    pass
            self.assertEqual(identify_product_directory(directory), "us")

    def test_unknown_or_partial_directory_is_left_unidentified(self):
        with tempfile.TemporaryDirectory() as directory:
            with open(os.path.join(directory, "btc_chain.xlsx"), "wb"):
                pass
            self.assertIsNone(identify_product_directory(directory))

    def test_invalid_marker_does_not_fall_back_to_workbooks(self):
        with tempfile.TemporaryDirectory() as directory:
            for filename in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
                with open(os.path.join(directory, filename), "wb"):
                    pass
            with open(os.path.join(directory, PRODUCT_MARKER), "w", encoding="ascii") as marker:
                marker.write("unknown")
            self.assertIsNone(identify_product_directory(directory))

    def test_cross_folder_delta_copy_replaces_legacy_delta_only(self):
        with tempfile.TemporaryDirectory() as first, tempfile.TemporaryDirectory() as legacy:
            for directory in (first, legacy):
                for filename in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
                    with open(os.path.join(directory, filename), "wb"):
                        pass
            self.assertTrue(same_product_directory(first, legacy))

    def test_delta_copy_does_not_replace_us_copy(self):
        with tempfile.TemporaryDirectory() as delta, tempfile.TemporaryDirectory() as us:
            for filename in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
                with open(os.path.join(delta, filename), "wb"):
                    pass
            for filename in (
                "spx_option_chain.xlsx",
                "djx_option_chain.xlsx",
                "us_stock_option_chain.xlsx",
            ):
                with open(os.path.join(us, filename), "wb"):
                    pass
            self.assertFalse(same_product_directory(delta, us))

    def test_unknown_directory_is_not_a_same_product_match(self):
        with tempfile.TemporaryDirectory() as delta, tempfile.TemporaryDirectory() as unknown:
            for filename in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
                with open(os.path.join(delta, filename), "wb"):
                    pass
            self.assertFalse(same_product_directory(delta, unknown))


if __name__ == "__main__":
    unittest.main()
