from __future__ import annotations

import sys
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parents[2] / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from build_workbooks import build_test_workbook  # noqa: E402
from pyvbaharness import ExcelSession  # noqa: E402


@pytest.fixture(scope="session")
def excel():
    workbook = build_test_workbook()
    with ExcelSession() as session:
        session.open_workbook(str(workbook))
        yield session


def _runner(excel, module: str):
    def run(proc: str, *args, timeout: int = 60):
        result = excel.run_macro(f"{module}.{proc}", *args, timeout=timeout)
        assert result.outcome == "passed", (
            f"{proc} outcome={result.outcome} error={result.error}"
        )
        return result.value

    return run


@pytest.fixture
def run_core(excel):
    return _runner(excel, "TestReDimCore")


@pytest.fixture
def run_async(excel):
    return _runner(excel, "TestReDimAsync")


def parse_transcript(raw: str) -> dict[str, str]:
    return dict(token.split("=", 1) for token in raw.split("|"))
