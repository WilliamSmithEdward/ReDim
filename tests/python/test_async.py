"""Live async engine tests: ops, click sugar, transports, cancel, jobs, timer."""

from __future__ import annotations

from conftest import parse_transcript


def test_async_op_lifecycle(run_async):
    facts = parse_transcript(run_async("TestAsyncOpLifecycle"))
    assert facts["statusAfterStart"] == "running"
    assert facts["busyText"] == "Working..."
    assert facts["spinnerShown"] == "True"
    assert facts["spinnerIsRing"] == "True", (
        "spinner must be a block arc; plain arcs wobble when rotated"
    )
    assert facts["workBeforeTick"] == "0"
    assert facts["clickWhileBusyIgnored"] == "True"
    assert facts["workAfterTick"] == "1"
    assert facts["busyDuringWork"] == "True"
    assert facts["doneRan"] == "1"
    assert facts["statusAfterTick"] == "done"
    assert facts["restoredText"] == "Run"
    assert facts["spinnerHidden"] == "True"
    assert facts["labelText"] == "Op: done"


def test_on_click_async_sugar(run_async):
    facts = parse_transcript(run_async("TestOnClickAsyncSugar"))
    assert facts["busyAfterClick"] == "True"
    assert facts["textAfterClick"] == "Working..."
    assert facts["workBeforeTick"] == "0", "clicks while busy must not queue work"
    assert facts["workAfterTick"] == "1"
    assert facts["busyAfterTick"] == "False"
    assert facts["restoredText"] == "Run"
    assert facts["secondRunWorks"] == "True"


def test_transport_op_across_ticks(run_async):
    facts = parse_transcript(run_async("TestTransportOpAcrossTicks"))
    assert facts["midStatus"] == "running", "delay must span multiple ticks"
    assert facts["finalStatus"] == "done"


def test_cancellation(run_async):
    facts = parse_transcript(run_async("TestCancellation"))
    assert facts["statusAfterStart"] == "running"
    assert facts["statusAfterCancel"] == "canceled"
    assert facts["workNeverRan"] == "True"
    assert facts["cancelRan"] == "1"
    assert facts["buttonRestored"] == "True"
    assert facts["midFlight"] == "running"
    assert facts["midFlightCanceled"] == "canceled"


def test_job_chunks(run_async):
    facts = parse_transcript(run_async("TestJobChunks"))
    assert facts["chunked"] == "True", (
        f"first tick ran {facts['firstTickSteps']} steps; the budget must "
        "split the job across ticks"
    )
    assert facts["finalSteps"] == "5"
    assert facts["doneRan"] == "1"


def test_job_cancel(run_async):
    facts = parse_transcript(run_async("TestJobCancel"))
    assert facts["cancelRan"] == "1"
    assert facts["stoppedEarly"] == "True"
    assert facts["stillRunning"] == "False"


def test_real_timer_end_to_end(run_async):
    facts = parse_transcript(run_async("TestRealTimerEndToEnd"))
    assert facts["armedAfterStart"] == "True"
    assert facts["cursorPinned"] == "True", (
        "an armed pump must pin the cursor so busy-cursor strobing cannot happen"
    )
    assert facts["statusNoManualPump"] == "done", (
        "the armed SetTimer alone must complete the op"
    )
    assert facts["autoDisarmed"] == "True", (
        "the pump must kill its own timer once work drains"
    )
    assert facts["cursorRestored"] == "True"
