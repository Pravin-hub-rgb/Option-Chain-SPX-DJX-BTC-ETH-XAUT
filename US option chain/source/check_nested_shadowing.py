"""Find names that a nested function assigns but that also exist at module
level. Inside a nested function, such an assignment creates a LOCAL unless
the name is declared global, so the module-level value silently never
changes.

Delta had one: the restore() closure assigned _last_bulk_restore, and the
exit path that reads it therefore never found a restore to call, so Excel
could be left in manual calculation after a crash.
"""

import ast
import io
import sys


def module_names(tree):
    names = set()
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name):
                    names.add(t.id)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
    return names


def nested_assignments(tree):
    """(function, name) for every assignment made inside a nested function."""
    found = []
    for outer in ast.walk(tree):
        if not isinstance(outer, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        for inner in ast.walk(outer):
            if inner is outer:
                continue
            if not isinstance(inner, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            declared = set()
            for stmt in ast.walk(inner):
                if isinstance(stmt, ast.Global):
                    declared.update(stmt.names)
            for node in ast.walk(inner):
                if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
                    if node.id not in declared:
                        found.append((outer.name, inner.name, node.id, node.lineno))
    return found


def main(path):
    with io.open(path, encoding="utf-8-sig", newline="") as handle:
        src = handle.read()
    tree = ast.parse(src)
    top = module_names(tree)
    problems = []
    for outer, inner, name, lineno in nested_assignments(tree):
        if name in top:
            problems.append((lineno, outer, inner, name))
    problems.sort()
    for lineno, outer, inner, name in problems:
        print(f"{path}:{lineno}: {outer} -> {inner} assigns {name!r}, "
              f"which also exists at module level")
    if not problems:
        print(f"{path}: no nested assignment shadows a module-level name")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
