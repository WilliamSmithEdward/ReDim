"""A workbook whose name holds an apostrophe, such as "Bob's Budget.xlsm":
ReDim draws in it, and the click macro it gives every shape resolves.

The test saves the built test workbook under such a name, so it runs in an
Excel session of its own: the shared session's workbook keeps its name.
"""

from __future__ import annotations

import sys
from pathlib import Path

TOOLS = Path(__file__).resolve().parents[2] / "tools"
sys.path.insert(0, str(TOOLS))

from build_workbooks import build_test_workbook  # noqa: E402
from pyvbaharness import ExcelSession  # noqa: E402


def test_apostrophe_in_workbook_name(tmp_path):
    workbook = build_test_workbook()
    renamed = tmp_path / "Bob's Tests.xlsm"
    with ExcelSession() as session:
        session.open_workbook(str(workbook))
        result = session.run_macro(
            "TestReDimWidgets.TestApostropheName", str(renamed), timeout=60
        )
    assert result.outcome == "passed", (
        f"outcome={result.outcome} error={result.error} message={result.message}"
    )
    facts = dict(pair.split("=", 1) for pair in str(result.value).split("|"))
    assert facts["name"] == "True", "the workbook runs under a name holding an apostrophe"
    assert facts["renders"] == "True", "ReDim draws in that workbook"
    assert facts["resolves"] == "True", "the click macro on its shapes resolves"
