"""Static checks on the Delta source.

Two failure classes are locked down here:

1. A bare name the module references but never binds. The US split turned
   several of these into runtime NameErrors that only the live pipeline hit.

2. A name a function reads before it has ever assigned it. Python makes a
   name local to a function as soon as the function assigns it anywhere, so
   reading it before the assignment raises UnboundLocalError even when the
   name is perfectly bound at module level. The module-level check cannot
   see this, because it collects every assignment in the file regardless of
   scope.

   Three real bugs were exactly this shape in the US tool, and one of them
   sat behind a flag nobody ever set, so nothing noticed until the code was
   read.
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


def _scope_bindings(node):
    """(assigned, loaded, declared, params) for one function scope.

    Nested functions, lambdas and classes have their own scope and are not
    descended into. A list comprehension has its own scope for its loop
    variable, but a walrus inside it binds in the enclosing scope.
    """
    assigned = {}
    loaded = []
    declared = set()

    def walk(current, in_comprehension, comprehension_names=frozenset()):
        if isinstance(current, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda, ast.ClassDef)):
            return
        if isinstance(current, (ast.ListComp, ast.SetComp, ast.GeneratorExp, ast.DictComp)):
            scoped = set(comprehension_names)
            for gen in current.generators:
                for child in ast.walk(gen.target):
                    if isinstance(child, ast.Name):
                        scoped.add(child.id)
            if isinstance(current, ast.DictComp):
                walk(current.key, True, scoped)
                walk(current.value, True, scoped)
            else:
                walk(current.elt, True, scoped)
            for gen in current.generators:
                walk(gen.iter, True, scoped)
                for cond in gen.ifs:
                    walk(cond, True, scoped)
            return
        if isinstance(current, (ast.Global, ast.Nonlocal)):
            declared.update(current.names)
            return
        if isinstance(current, ast.Name):
            if isinstance(current.ctx, ast.Store):
                if not in_comprehension:
                    assigned.setdefault(current.id, current.lineno)
            elif isinstance(current.ctx, ast.Load):
                if not (in_comprehension and current.id in comprehension_names):
                    loaded.append((current.lineno, current.id))
            return
        if isinstance(current, ast.NamedExpr):
            target = current.target
            if isinstance(target, ast.Name):
                assigned.setdefault(target.id, target.lineno)
            return
        if isinstance(current, ast.ExceptHandler):
            if current.name:
                assigned.setdefault(current.name, current.lineno)
        if isinstance(current, (ast.Import, ast.ImportFrom)):
            for alias in current.names:
                name = alias.asname or alias.name.split(".")[0]
                assigned.setdefault(name, current.lineno)
            return
        for child in ast.iter_child_nodes(current):
            walk(child, in_comprehension, comprehension_names)

    args = node.args
    params = set()
    for arg in list(args.args) + list(args.posonlyargs) + list(args.kwonlyargs):
        params.add(arg.arg)
    if args.vararg:
        params.add(args.vararg.arg)
    if args.kwarg:
        params.add(args.kwarg.arg)

    for statement in node.body:
        walk(statement, False)
    return assigned, loaded, declared, params


def unbound_locals(path):
    """Names a function reads before it has ever assigned them."""
    with io.open(path, encoding="utf-8-sig", newline="") as handle:
        source = handle.read()
    tree = ast.parse(source)
    flagged = []
    for node in ast.walk(tree):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        assigned, loaded, declared, params = _scope_bindings(node)
        for lineno, name in loaded:
            if name in declared or name in params or name not in assigned:
                continue
            if lineno < assigned[name]:
                flagged.append((node.name, lineno, name))
    return flagged


class NameBindingTests(unittest.TestCase):
    RUNNABLE = [
        "main.py",
        "calc_mode_guard.py",
        "config.py",
        "delta_diagnostics.py",
        "delta_instance.py",
        "delta_licensing.py",
        "delta_performance.py",
        "delta_state.py",
        "product_identity.py",
    ]

    def test_runnable_modules_have_no_unbound_names(self):
        for fname in self.RUNNABLE:
            with self.subTest(module=fname):
                flagged = undefined_references(os.path.join(SOURCE_DIR, fname))
                self.assertEqual([], flagged)

    def test_runnable_modules_have_no_names_read_before_they_are_assigned(self):
        for fname in self.RUNNABLE:
            with self.subTest(module=fname):
                flagged = unbound_locals(os.path.join(SOURCE_DIR, fname))
                self.assertEqual([], flagged)


class NestedShadowingTests(unittest.TestCase):
    """A nested function assigning a module-level name creates a local.

    Inside a nested function, an assignment to a name that also exists at
    module level is a LOCAL unless the name is declared global. The module
    level value then silently never changes.

    Delta had one: the restore() closure assigned _last_bulk_restore, so the
    exit path that reads it never found a restore to call, and Excel could be
    left in manual calculation after a crash.
    """

    def test_no_nested_assignment_shadows_a_module_level_name(self):
        import check_nested_shadowing

        for fname in NameBindingTests.RUNNABLE:
            with self.subTest(module=fname):
                path = os.path.join(SOURCE_DIR, fname)
                with io.open(path, encoding="utf-8-sig", newline="") as handle:
                    src = handle.read()
                tree = ast.parse(src)
                top = check_nested_shadowing.module_names(tree)
                problems = []
                for outer, inner, name, lineno in check_nested_shadowing.nested_assignments(tree):
                    if name in top:
                        problems.append(f"{fname}:{lineno} {outer}->{inner} assigns {name}")
                self.assertEqual([], problems)


if __name__ == "__main__":
    unittest.main()
