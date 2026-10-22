"""The A-split turned bare names into runtime NameErrors that only the live
pipeline hits (spot_source_for, refresh_tencent_spot_cache, SPOT_POLL_SECONDS).
Lock that failure class behind a static check: every bare name the modules
reference must be bound in that module.
"""

import ast
import builtins
import io
import os
import unittest

SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))

BUILTINS = set(dir(builtins)) | {"__file__", "__name__", "self"}


def _add_targets(names, target):
    if target is not None:
        names.update(
            t.id for t in ast.walk(target) if isinstance(t, ast.Name)
        )


def _add_params(names, node):
    a = node.args
    for arg in list(a.args) + list(a.posonlyargs) + list(a.kwonlyargs):
        names.add(arg.arg)
    if a.vararg:
        names.add(a.vararg.arg)
    if a.kwarg:
        names.add(a.kwarg.arg)


def collect_bindings(tree):
    names = set(BUILTINS)
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            names.add(node.name)
            _add_params(names, node)
        elif isinstance(node, ast.Lambda):
            _add_params(names, node)
        elif isinstance(node, ast.ClassDef):
            names.add(node.name)
        elif isinstance(node, ast.Import):
            for alias in node.names:
                names.add((alias.asname or alias.name).split(".")[0])
        elif isinstance(node, ast.ImportFrom):
            for alias in node.names:
                names.add(alias.asname or alias.name)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                _add_targets(names, target)
        elif isinstance(node, ast.AnnAssign):
            _add_targets(names, node.target)
        elif isinstance(node, ast.AugAssign):
            _add_targets(names, node.target)
        elif isinstance(node, ast.For):
            _add_targets(names, node.target)
        elif isinstance(node, (ast.With, ast.AsyncWith)):
            for item in node.items:
                _add_targets(names, item.optional_vars)
        elif isinstance(node, ast.comprehension):
            _add_targets(names, node.target)
        elif isinstance(node, ast.ExceptHandler):
            if node.name:
                names.add(node.name)
        elif isinstance(node, ast.NamedExpr):
            names.add(node.target.id)
    return names


def undefined_references(path):
    with io.open(path, encoding="utf-8-sig", newline="") as handle:
        source = handle.read()
    tree = ast.parse(source)
    bindings = collect_bindings(tree)
    flagged = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
            if node.id not in bindings:
                flagged.append((node.lineno, node.id))
    return flagged


class NameBindingTests(unittest.TestCase):
    def test_runnable_modules_have_no_unbound_names(self):
        runnable = [
            "main.py",
            "us_chain.py",
            "us_excel.py",
            "us_instance.py",
            "us_licensing.py",
            "us_performance.py",
            "us_spot.py",
            "us_state.py",
            "us_time.py",
            "us_workbook.py",
            "us_diagnostics.py",
            "excel_helpers.py",
            "calc_mode_guard.py",
            "config.py",
            "product_identity.py",
            "databento_feed.py",
        ]
        for fname in runnable:
            with self.subTest(module=fname):
                flagged = undefined_references(os.path.join(SOURCE_DIR, fname))
                self.assertEqual([], flagged)


if __name__ == "__main__":
    unittest.main()
