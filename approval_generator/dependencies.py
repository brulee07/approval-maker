"""Check all runtime imports before opening the desktop application."""
from __future__ import annotations

import importlib
import sys


# Import name -> distribution name; tkinter is supplied by the Python installer.
REQUIRED_MODULES = {
    "tkinter": None,
    "openpyxl": "openpyxl",
    "pdfplumber": "pdfplumber",
    "pypdf": "pypdf",
    "reportlab": "reportlab",
    "pypdfium2": "pypdfium2",
    "PIL": "Pillow",
    "lxml": "lxml",
    "xlrd": "xlrd",
}


def unavailable_modules() -> list[str]:
    missing = []
    for module in REQUIRED_MODULES:
        try:
            importlib.import_module(module)
        except (ImportError, OSError):
            missing.append(module)
    return missing


def main() -> int:
    if sys.version_info < (3, 11):
        print("Python 3.11 or newer is required. Please install Python from python.org.")
        return 2
    missing = unavailable_modules()
    if "tkinter" in missing:
        print("Tcl/Tk is missing. Please reinstall Python with Tcl/Tk support, then run start.bat.")
        return 2
    if missing:
        print("Required components are missing: " + ", ".join(missing))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
