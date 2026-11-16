import contextlib
import io
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch

from approval_generator import dependencies
from approval_generator.extract import extract_quote
from approval_generator.models import UserError


class DependencyTests(unittest.TestCase):
    def test_every_distributed_requirement_is_checked(self):
        requirements = Path(__file__).resolve().parent.parent / "requirements.txt"
        packages = {re.split(r"[<>=!~\[]", line)[0].strip().lower()
                    for line in requirements.read_text(encoding="utf-8").splitlines()
                    if line.strip() and not line.startswith("#")}
        checked = {name.lower() for name in dependencies.REQUIRED_MODULES.values() if name}
        self.assertEqual(packages, checked)

    def test_missing_xlrd_is_found_before_the_gui_opens(self):
        def imported(module):
            if module == "xlrd":
                raise ModuleNotFoundError("No module named 'xlrd'")
            return object()
        with patch.object(dependencies.importlib, "import_module", side_effect=imported):
            self.assertEqual(dependencies.unavailable_modules(), ["xlrd"])
            with contextlib.redirect_stdout(io.StringIO()) as messages:
                self.assertEqual(dependencies.main(), 1)
            self.assertIn("xlrd", messages.getvalue())

    def test_complete_environment_does_not_request_installation(self):
        with patch.object(dependencies.importlib, "import_module", return_value=object()):
            self.assertEqual(dependencies.main(), 0)

    def test_missing_tk_requires_python_repair_instead_of_pip(self):
        with patch.object(dependencies, "unavailable_modules", return_value=["tkinter", "xlrd"]), \
             contextlib.redirect_stdout(io.StringIO()) as messages:
            self.assertEqual(dependencies.main(), 2)
        self.assertIn("Tcl/Tk", messages.getvalue())

    def test_xls_error_names_the_shipped_installer(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "quotation.xls"
            source.write_bytes(b"placeholder")
            with patch.dict("sys.modules", {"xlrd": None}), self.assertRaises(UserError) as error:
                extract_quote(str(source))
            self.assertIn("install.bat", str(error.exception))
            self.assertNotIn("설치.bat", str(error.exception))


if __name__ == "__main__":
    unittest.main()
