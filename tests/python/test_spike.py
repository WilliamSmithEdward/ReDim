"""Architecture spike: prove the pump model under a live Excel session.

Verified assumptions:

1. A SetTimer callback registered from workbook VBA fires whenever Excel pumps
   messages, and ROneCOne's Friend AdvanceTask steps tasks from that callback.
2. A task nobody awaits reaches RanToCompletion purely through pump ticks.
3. Shapes can be created, styled, and dispatched through OnAction.

Also pinned here: a harness constraint discovered during the spike. Module
globals do not survive across run_macro round trips, so every live scenario
must complete inside a single VBA call, and a SetTimer must never stay armed
across harness call boundaries (its TIMERPROC address goes stale with the
project reset and the next WM_TIMER kills the process). Production Excel does
not mutate the project between clicks, so this rule is test-only; the
framework still ships orphan-timer recovery for real state-loss events.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parents[2] / "tools"
sys.path.insert(0, str(TOOLS))

from vba_sources import (  # noqa: E402
    OUTPUT,
    TESTS_VBA,
    prepare_class_source,
    read_vba,
    ronecone_class_path,
)

from pyopenvba import ExcelFile, VBAModuleKind  # noqa: E402
from pyvbaharness import ExcelSession  # noqa: E402

SPIKE_WORKBOOK = OUTPUT / "spike.xlsm"


def build_spike_workbook() -> Path:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    runtime_source = prepare_class_source(ronecone_class_path())
    spike_source = read_vba(TESTS_VBA / "SpikeHost.bas")
    if SPIKE_WORKBOOK.exists():
        SPIKE_WORKBOOK.unlink()
    with ExcelFile.create_new(SPIKE_WORKBOOK) as workbook:
        project = workbook.vba_project()
        project.add_module("ROneCOne", runtime_source, kind=VBAModuleKind.other)
        project.add_module("SpikeHost", spike_source, kind=VBAModuleKind.standard)
        workbook.save()
    return SPIKE_WORKBOOK


@pytest.fixture(scope="module")
def excel():
    workbook = build_spike_workbook()
    with ExcelSession() as session:
        session.open_workbook(str(workbook))
        yield session


def run(excel, proc, *args, timeout=30):
    result = excel.run_macro(f"SpikeHost.{proc}", *args, timeout=timeout)
    assert result.outcome == "passed", (
        f"{proc} outcome={result.outcome} error={result.error}"
    )
    return result.value


def test_globals_do_not_survive_run_macro(excel):
    """Pins the harness constraint the live-test design is built around."""
    assert run(excel, "SetGlobal", 42) == 42
    assert run(excel, "GetGlobal") == 0


def test_pump_fires_while_vba_pumps_messages(excel):
    ticks, errors = run(excel, "PumpProofSingleCall").split("|")
    assert int(ticks) >= 3, f"expected pump ticks, got {ticks}"
    assert int(errors) == 0


def test_pump_completes_unawaited_task(excel):
    status, done_tick, ticks, errors = run(
        excel, "TaskPumpProofSingleCall"
    ).split("|")
    assert status == "RanToCompletion"
    assert int(done_tick) > 0
    assert int(ticks) > int(done_tick)
    assert int(errors) == 0


def test_shapes_create_read_rotate_dispatch(excel):
    assert run(excel, "CreateSpikeShape") == "rdm_spike_button"
    name, text, on_action = run(excel, "ReadSpikeShape").split("|")
    assert name == "rdm_spike_button"
    assert text == "Spike"
    assert on_action == "SpikeHost.HandleSpikeClick"
    assert run(excel, "RotateSpikeShape") == 30.0
    assert run(excel, "DispatchByName", "rdm_spike_button") == "rdm_spike_button"
