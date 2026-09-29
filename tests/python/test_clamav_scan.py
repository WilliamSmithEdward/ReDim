"""ClamAV report parsing and reviewed finding behavior, without a scanner install."""

from pathlib import Path
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import security_scan  # noqa: E402


class ClamAVTests(unittest.TestCase):
    def test_parses_windows_paths_and_errors(self):
        output = ("C:\\scan\\demo.xlsm: Macro.Test FOUND\n"
                  "C:\\scan\\broken.xlsm: Permission denied ERROR\n")
        found, errors = security_scan.parse_clamav_output(
            output, {"C:\\scan\\demo.xlsm": "file:demo.xlsm"})
        self.assertEqual(found, [("file:demo.xlsm", "Macro.Test")])
        self.assertEqual(len(errors), 1)

    def test_reviewed_signature_is_scoped_to_its_module(self):
        module = security_scan.Module("ROneCOne", "WScript.Shell", "test")
        expected = {"reasons": {"known": "Reviewed process surface."},
                    "clamav": {"module:ROneCOne": {"Macro.Test": "known"}}}

        def fake_run(command, **_kwargs):
            if "--version" in command:
                return security_scan.subprocess.CompletedProcess(command, 0, "ClamAV test/1", "")
            module_path = next(path for path in command if str(path).endswith(".vba"))
            output = f"{module_path}: Macro.Test FOUND\n"
            return security_scan.subprocess.CompletedProcess(command, 1, output, "")

        with patch.object(security_scan.subprocess, "run", side_effect=fake_run):
            scan = security_scan.scan_clamav([], [module], expected)
        self.assertEqual(scan.findings, [
            ("module:ROneCOne", "Macro.Test", "Reviewed process surface.")])
        self.assertEqual(scan.errors, [])

        with patch.object(security_scan.subprocess, "run", side_effect=fake_run):
            unreviewed = security_scan.scan_clamav([], [module],
                                                  {"reasons": expected["reasons"], "clamav": {}})
        self.assertEqual(unreviewed.findings[0][2], None)

    def test_scanner_error_is_not_treated_as_clean(self):
        def fake_run(command, **_kwargs):
            if "--version" in command:
                return security_scan.subprocess.CompletedProcess(command, 0, "ClamAV test/1", "")
            return security_scan.subprocess.CompletedProcess(command, 2, "", "database missing")

        with patch.object(security_scan.subprocess, "run", side_effect=fake_run):
            scan = security_scan.scan_clamav([], [], {"reasons": {}, "clamav": {}})
        self.assertTrue(scan.errors)


if __name__ == "__main__":
    unittest.main()
