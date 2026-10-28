"""Regression test: the frozen US build ships only the pyarrow closure it uses.

import databento pulls pyarrow unconditionally (databento/__init__ imports
databento.common.dbnstore, which does `import pyarrow` + `import pyarrow.parquet`
at module top), so pyarrow CANNOT be excluded wholesale. The verified pull chain
on 2026-10-25 (sys.modules after `import databento`) needs only
pyarrow.lib/_compute/ipc/types/util + pyarrow.parquet + pyarrow.fs and the cloud
filesystem backends pyarrow.fs/__init__ imports eagerly
(_s3fs/_gcsfs/_azurefs/_hdfs).

build_exe.bat excludes flight/substrait/dataset/orc/csv/json/feather/acero/
parquet.encryption and trims their hook-shipped .pyd/.dll/pyx payloads. This test
pins the trimmed bundle: the excluded submodules must stay out, and every binary
the pull chain needs must still load with its PE import closure intact.
"""

import ctypes
import glob
import os
import unittest

_DIST_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist",
                          "us_option_chain")
_PYARROW = os.path.join(_DIST_ROOT, "_internal", "pyarrow")

_EXCLUDED_PREFIXES = (
    "_acero", "_csv", "_dataset", "_feather", "_flight", "_json", "_orc",
    "_parquet_encryption", "_substrait",
)
_EXCLUDED_DLLS = (
    "arrow_acero.dll", "arrow_dataset.dll", "arrow_flight.dll",
    "arrow_python_flight.dll", "arrow_python_parquet_encryption.dll",
    "arrow_substrait.dll",
)
_REQUIRED_FILES = (
    "arrow.dll", "arrow_compute.dll", "arrow_python.dll", "parquet.dll",
    "_compute.cp310-win_amd64.pyd", "_fs.cp310-win_amd64.pyd",
    "_parquet.cp310-win_amd64.pyd", "_s3fs.cp310-win_amd64.pyd",
)


def _pyd(prefix):
    """The frozen pyd carries its ABI tag (_acero.cp310-win_amd64.pyd)."""
    hits = glob.glob(os.path.join(_PYARROW, prefix + ".*.pyd"))
    return os.path.basename(hits[0]) if hits else None


class USPyArrowPayloadTests(unittest.TestCase):
    """Trim forever, but never so far that the databento import chain breaks."""

    def setUp(self):
        if not os.path.isdir(_PYARROW):
            self.skipTest("no frozen dist build present - build_exe.bat first")

    def test_excluded_submodules_are_not_shipped(self):
        shipped = os.listdir(_PYARROW)
        for name in shipped:
            for prefix in _EXCLUDED_PREFIXES:
                # the dot keeps _parquet_encryption/_dataset_orc distinct from
                # _parquet/_dataset; prefix+"." cannot match a same-prefix file
                if name.startswith(prefix + "."):
                    self.fail(f"excluded payload shipped: {name}")
        for dll in _EXCLUDED_DLLS:
            self.assertNotIn(dll, shipped, f"excluded DLL shipped: {dll}")

    def test_pyarrow_cython_sources_and_dev_tree_are_not_shipped(self):
        shipped = os.listdir(_PYARROW)
        for name in shipped:
            self.assertFalse(name.endswith(".pyx"), f"pyx shipped: {name}")
            self.assertFalse(name.endswith(".pxi"), f"pxi shipped: {name}")
        for tree in ("tests", "src", "includes"):
            self.assertNotIn(tree, shipped, f"dev tree shipped: {tree}/")

    def test_every_binary_the_databento_pull_chain_needs_survives(self):
        for name in _REQUIRED_FILES:
            self.assertTrue(os.path.isfile(os.path.join(_PYARROW, name)),
                            f"required pyarrow binary missing: {name}")
        for backend in ("_s3fs", "_gcsfs", "_azurefs", "_hdfs"):
            pyd = _pyd(backend)
            self.assertIsNotNone(
                pyd, f"cloud filesystem backend missing: {backend}")

    def test_kept_binaries_still_load_with_their_pe_import_closure(self):
        """LoadLibrary proves no deleted DLL was imported by a kept binary."""
        os.add_dll_directory(_PYARROW)
        pyarrow_libs = os.path.join(os.path.dirname(_PYARROW), "pyarrow.libs")
        if os.path.isdir(pyarrow_libs):
            os.add_dll_directory(pyarrow_libs)
        wanted = list(_REQUIRED_FILES)
        for backend in ("_s3fs", "_gcsfs", "_azurefs", "_hdfs"):
            pyd = _pyd(backend)
            if pyd:
                wanted.append(pyd)
        for name in wanted:
            try:
                ctypes.WinDLL(os.path.join(_PYARROW, name))
            except OSError as exc:
                self.fail(f"frozen binary no longer loads: {name}: {exc}")


if __name__ == "__main__":
    unittest.main()
