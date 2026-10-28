"""Report names that main.py imports but never uses.

Purely informational: it is a leftover list from the module split, not a
failure. The static name check in tests/test_delta_names.py is the one that
actually gates the build.
"""

import ast
import io
import sys


def main(path):
    src = io.open(path, encoding="utf-8-sig", newline="").read()
    tree = ast.parse(src)
    imported = {}
    for node in ast.walk(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            for alias in node.names:
                name = alias.asname or alias.name.split(".")[0]
                imported[name] = node.lineno
    used = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
            used.add(node.id)
        elif isinstance(node, ast.Attribute):
            pass
    # Attribute bases are Names too, so they are already collected.
    unused = sorted(n for n in imported if n not in used and n != "*")
    for name in unused:
        print(f"{path}:{imported[name]}: {name} imported but never used")
    if not unused:
        print(f"{path}: no unused imports")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
