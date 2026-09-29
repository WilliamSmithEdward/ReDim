"""tools/security_scan.py's own logic, without oletools or Excel: the p-code
reading that decides a workbook holds code its source does not show, the
module text it compares, and the expected-findings file it judges by."""

from __future__ import annotations

import json

import security_scan

# pcodedmp's report on a v1.0.4 workbook: pyOpenVBA's template left Module1
# with no source and the p-code of a comment; the modules ReDim injected
# hold source alone, which pcodedmp misreads as p-code and reports as an
# error.
V104_PCODEDMP = """\
Identifiers:

Module streams:
VBA/ThisWorkbook - 1188 bytes
VBA/Sheet1 - 991 bytes
VBA/Module1 - 749 bytes
Line #0:
\tQuoteRem 0x0000 0x0030 "TESTING ONLY DO NOT INCLUDE THIS IN FINAL OUTPUT"
VBA/ROneCOne - 355548 bytes
Error: unpack_from requires a buffer of at least 1835093615 bytes for unpacking 4 bytes \
at offset 1835093611 (actual buffer size is 355548).
VBA/ReDimUI - 410914 bytes
Error: unpack_from requires a buffer of at least 1835093615 bytes for unpacking 4 bytes \
at offset 1835093611 (actual buffer size is 410914).
"""


def test_pcode_lines_finds_the_template_comment():
    assert security_scan.pcode_lines(V104_PCODEDMP) == {
        "Module1": ['QuoteRem 0x0000 0x0030 "TESTING ONLY DO NOT INCLUDE THIS IN FINAL OUTPUT"'],
    }


def test_pcode_lines_passes_source_only_modules():
    without_module1 = V104_PCODEDMP.replace(
        'VBA/Module1 - 749 bytes\nLine #0:\n\tQuoteRem 0x0000 0x0030 '
        '"TESTING ONLY DO NOT INCLUDE THIS IN FINAL OUTPUT"\n', "")
    assert "Module1" not in without_module1
    assert security_scan.pcode_lines(without_module1) == {}


def test_module_code_matches_a_class_file_to_its_workbook_module():
    exported = ('VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                'Attribute VB_Name = "Sample"\r\nOption Explicit\r\n')
    stored = 'Attribute VB_Name = "Sample"\nOption Explicit\n'
    assert security_scan.module_code(exported) == stored
    assert security_scan.module_code(stored) == stored


def test_short_encoded_rule():
    for text in ("2147483647", "LEFT", "Alt+", "SHA1"):
        assert security_scan.SHORT_ENCODED.fullmatch(text), text
    assert not security_scan.SHORT_ENCODED.fullmatch("cG93ZXJzaGVsbCAtZW5j")
    assert not security_scan.SHORT_ENCODED.fullmatch("C:\\Temp")


def test_every_expected_finding_names_a_reason():
    expected = json.loads(security_scan.EXPECTED_PATH.read_text(encoding="utf-8"))
    reasons = expected["reasons"]
    for module, entries in expected["modules"].items():
        for key, reason in entries.items():
            assert reason in reasons, f"{module} {key} names no reason: {reason}"
            assert ": " in key, f"{module} {key} is not 'Type: keyword'"
    used = {reason for entries in expected["modules"].values() for reason in entries.values()}
    assert "encoded" in reasons
    assert set(reasons) - used - {"encoded"} == set(), "reasons no entry uses"
