"""The upstream bundle must be exact, bounded, and verified before scanning."""

import hashlib
import importlib.util
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import yara_forge  # noqa: E402
import security_scan  # noqa: E402


def archive(member: str = yara_forge.MEMBER, data: bytes = b"rule test { condition: true }") -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr(member, data)
    return buffer.getvalue()


class YaraForgeTests(unittest.TestCase):
    def test_archive_layout_is_exact(self):
        self.assertIn(b"rule test", yara_forge.unpack(archive()))
        with self.assertRaises(ValueError):
            yara_forge.unpack(archive("other.yar"))

    def test_fetch_rejects_changed_archive(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pin = root / "pin.json"
            pin.write_text(json.dumps({"yara_forge": {"release": "20260927", "asset": yara_forge.ASSET,
                                       "url": yara_forge.release_asset("20260927"),
                                       "sha256": "0" * 64}}), encoding="utf-8")
            with patch.object(yara_forge, "PIN", pin), patch.object(yara_forge, "RULES", root / "core.yar"), \
                 patch.object(yara_forge, "download", return_value=archive()):
                with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
                    yara_forge.fetch()
                self.assertFalse((root / "core.yar").exists())

    def test_fetch_writes_only_verified_rule(self):
        data = archive()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pin = root / "pin.json"
            pin.write_text(json.dumps({"yara_forge": {"release": "20260927", "asset": yara_forge.ASSET,
                                       "url": yara_forge.release_asset("20260927"),
                                       "sha256": hashlib.sha256(data).hexdigest()}}), encoding="utf-8")
            rules = root / "core.yar"
            with patch.object(yara_forge, "PIN", pin), patch.object(yara_forge, "RULES", rules), \
                 patch.object(yara_forge, "download", return_value=data):
                yara_forge.fetch()
            self.assertEqual(rules.read_bytes(), yara_forge.unpack(data))

    def test_the_pin_file_is_one_the_standard_updater_accepts(self):
        root = Path(__file__).resolve().parents[2]
        spec = importlib.util.spec_from_file_location(
            "yara_update", root / ".github" / "security" / "yara_update.py")
        updater = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(updater)
        pins = json.loads((root / ".github" / "security" / "yara.json").read_text(encoding="utf-8"))
        updater.check_move(pins, pins)
        self.assertEqual(yara_forge.read_pin(), pins["yara_forge"])

    def test_forge_match_requires_reviewed_file_or_module_reason(self):
        try:
            import yara_x  # noqa: F401
        except ImportError:
            self.skipTest("YARA-X is not installed")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "sample.bas"
            source.write_text("test indicator", encoding="utf-8")
            rules = root / "core.yar"
            rules.write_text('rule Demo_Finding { strings: $a = "test indicator" condition: $a }',
                             encoding="utf-8")
            file = security_scan.ScannedFile(source, "", 14)
            module = security_scan.Module("Sample", "test indicator", "")
            expected = {"reasons": {"known": "Reviewed fixture"}, "forge": {
                "file:sample.bas": {"Demo_Finding": "known"},
                "module:Sample": {"Demo_Finding": "known"}}}
            findings, problems = security_scan.forge_findings([file], [module], rules, expected)
            self.assertEqual(len(findings), 2)
            self.assertFalse(problems)
            expected["forge"]["module:Sample"] = {}
            _, problems = security_scan.forge_findings([file], [module], rules, expected)
            self.assertEqual(problems, ["module:Sample: unexpected YARA Forge rule Demo_Finding matched"])


if __name__ == "__main__":
    unittest.main()
