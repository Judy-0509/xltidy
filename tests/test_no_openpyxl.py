import importlib
import sys


def test_importing_xltidy_does_not_load_openpyxl():
    for mod in [m for m in sys.modules if m == "openpyxl" or m.startswith("openpyxl.")]:
        del sys.modules[mod]
    importlib.import_module("xltidy")
    for name in ["coords", "models", "encode", "spec", "reconcile", "apply",
                 "consolidate", "dbio", "config", "infer"]:
        try:
            importlib.import_module(f"xltidy.{name}")
        except ModuleNotFoundError:
            pass
    assert "openpyxl" not in sys.modules, "openpyxl is a banned dependency (hard policy)"


def test_no_banned_excel_io_in_source():
    import pathlib
    src = pathlib.Path(__file__).resolve().parents[1] / "src" / "xltidy"
    offenders = []
    for p in src.rglob("*.py"):
        text = p.read_text(encoding="utf-8")
        for needle in ("read_excel", "ExcelFile", "import openpyxl", "from openpyxl"):
            if needle in text:
                offenders.append(f"{p.name}: {needle}")
    assert not offenders, f"banned excel I/O found: {offenders}"
