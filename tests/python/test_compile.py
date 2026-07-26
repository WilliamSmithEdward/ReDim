"""Compile gate: the real VBA compiler accepts every shipped workbook.

pyvbaanalysis catches most problems without Excel, but a few grammar rules
(such as statement-position calls with multiple parenthesized arguments) only
the real compiler enforces. This gate closes that gap for the framework, the
test modules, and every demo.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parents[2] / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from build_workbooks import build_demo_workbooks, build_test_workbook  # noqa: E402
from pyvbaharness import ExcelSession  # noqa: E402


def workbook_paths():
    paths = [build_test_workbook()]
    paths.extend(build_demo_workbooks())
    return paths


@pytest.mark.parametrize("workbook", workbook_paths(), ids=lambda p: p.name)
def test_project_compiles(workbook):
    with ExcelSession() as excel:
        excel.open_workbook(str(workbook))
        result = excel.compile_project(watch_seconds=20)
        dialog = getattr(result, "dialog", None)
        message = getattr(dialog, "message", None) if dialog else None
        assert result.outcome == "accepted", f"{workbook.name}: {message}"
