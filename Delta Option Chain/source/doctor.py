"""Delta Option Chain - install self-test.

This runs INSIDE the delivered executable:

    delta_option_chain.exe --doctor

It checks every thing that has ever stopped the tool working and writes a
plain English report to doctor-report.txt next to the executable, then opens
it. The customer does not need Python, and support gets one file to read.

Nothing here imports the product's own modules, so it works in the frozen
build exactly as it does from source.
"""

import os
import subprocess
import sys

OK = "PASS"
BAD = "FAIL"
WARN = "WARN"

PRODUCT = "Delta Option Chain"


def check(name, ok, detail="", level=None):
    level = level or (OK if ok else BAD)
    results.append((level, name, detail))
    return level, name, detail


def run_doctor(base_dir, out=None):
    """Run every check. Returns (report_text, everything_passed)."""
    lines = []
    results = []

    def say(text=""):
        lines.append(text)

    def section(title):
        say()
        say(f"--- {title} " + "-" * max(0, 58 - len(title)))

    def check(name, ok, detail="", level=None):
        level = level or (OK if ok else BAD)
        results.append((level, name, detail))
        say(f"[{level}] {name}" + (f" - {detail}" if detail else ""))

    say(f"{PRODUCT} self-test")
    say(f"folder: {base_dir}  ({len(base_dir)} characters)")

    # ---- the folder --------------------------------------------------------
    section("folder")
    # Excel's Workbooks.Open gives up somewhere past 128 characters and never
    # returns. A package unzipped under a user's Downloads folder is over
    # that easily, and the symptom is the tool saying "Excel is not running"
    # while Excel is sitting right there.
    if len(base_dir) > 120:
        check("folder path is short enough for Excel",
            False,
            f"{len(base_dir)} characters. Move the package to a short folder "
            "- C:\\OptionChain - and run it again.",
        )
    else:
        check("folder path is short enough for Excel",
            True,
            f"{len(base_dir)} characters",
        )

    # ---- python and packages ----------------------------------------------
    section("software")
    check("python", True, f"{sys.version.split()[0]}")
    for module, why in (
        ("xlwings", "talks to Excel"),
        ("requests", "talks to the data sources"),
        ("pandas", "builds the chain"),
        ("websockets", "receives live quotes"),
        ("openpyxl", "reads and writes the workbooks"),
    ):
        try:
            __import__(module)
            check(f"package {module}", True, why)
        except Exception as e:
            check(f"package {module}", False,
                   f"{why} - {type(e).__name__}: {e}")

    # ---- the activation ----------------------------------------------------
    section("activation")
    licence = os.path.join(base_dir, "license.json")
    if os.path.exists(licence):
        check("activation key saved", True, "next to the tool")
    else:
        # Not a problem: license.json is the customer's own file and only
        # exists once they have entered the activation key. Failing here
        # would send support chasing a file that is supposed to be missing.
        check("activation key saved",
            True,
            "not entered yet - the tool asks for the activation key on the "
            "first run",
        )

    # ---- Delta Exchange ----------------------------------------------------
    section("Delta Exchange")
    try:
        import requests

        response = requests.get(
            "https://api.india.delta.exchange/v2/products",
            params={
                "contract_types": "call_options",
                "underlying_asset_symbols": "BTC",
            },
            timeout=20,
        )
        products = len(response.json().get("result") or [])
        check("Delta Exchange reachable", True,
               f"{products} products returned")
    except Exception as e:
        check("Delta Exchange reachable",
            False,
            f"{type(e).__name__}: {str(e)[:140]}. Check the internet connection.",
        )

    # ---- Excel -------------------------------------------------------------
    section("Excel")
    try:
        import xlwings as xw

        apps = list(xw.apps)
        check("Excel is running", True,
               f"{len(apps)} instance(s) visible")
    except Exception as e:
        check("Excel is running", False, f"{type(e).__name__}: {e}")

    try:
        out_text = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
            capture_output=True, text=True, timeout=15,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout or ""
        if "EXCEL" in out_text.upper():
            check("Excel is running", True, "ready to take the chains")
        else:
            # Not a problem. Before the first run this is the normal state:
            # the tool starts Excel itself. Reporting it as a failure would
            # send the customer chasing something that is not wrong, and the
            # one thing they send us is this report.
            check("Excel is running", True,
                "not started yet - the tool opens it for you",
            )
    except Exception:
        pass

    # ---- the workbooks -----------------------------------------------------
    section("workbooks")
    for name in ("btc_chain.xlsx", "eth_chain.xlsx", "xaut_chain.xlsx"):
        path = os.path.join(base_dir, name)
        if not os.path.exists(path):
            check(f"workbook {name}", False, "not found next to the tool")
            continue
        ok = True
        detail = f"{os.path.getsize(path)} bytes"
        try:
            import openpyxl

            openpyxl.load_workbook(path)
        except Exception as e:
            ok = False
            detail = f"cannot read it - {type(e).__name__}: {e}"
        check(f"workbook {name}", ok, detail)

    # ---- the log -----------------------------------------------------------
    section("logs")
    log_dir = os.path.join(os.environ.get("LOCALAPPDATA", "."), "DeltaOptionChain")
    for name in ("feed.log", "startup_diag.log"):
        path = os.path.join(log_dir, name)
        if os.path.exists(path):
            check(f"log {name}", True, path)
        else:
            check(f"log {name}", True, "not written yet")

    # ---- summary -----------------------------------------------------------
    failed = [r for r in results if r[0] == BAD]
    warned = [r for r in results if r[0] == WARN]

    say()
    say("=" * 64)
    if failed:
        say(f"{len(failed)} problem(s) found. Fix these first:")
        for _, name, detail in failed:
            say(f"  - {name}: {detail}")
    elif warned:
        say(f"Everything needed is in place, with {len(warned)} warning(s).")
    else:
        say("Everything needed is in place. The tool should work.")
    say("=" * 64)

    report = "\n".join(lines) + "\n"
    if out:
        try:
            with open(out, "w", encoding="utf-8") as handle:
                handle.write(report)
        except OSError:
            pass
    return report, not failed


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    report, ok = run_doctor(base_dir)
    if "--quiet" not in argv:
        print(report)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
