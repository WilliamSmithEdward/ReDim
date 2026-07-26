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


def test_adaptive_budget(run_async):
    facts = parse_transcript(run_async("TestAdaptiveBudget"))
    assert facts["yieldsToAnimation"] == "True", (
        "a default-budget job must take fewer steps per pass while animating"
    )
    assert int(facts["stepsAnimating"]) >= 1


def test_render_arms_drag_watch(run_async):
    facts = parse_transcript(run_async("TestRenderArmsDragWatch"))
    assert facts["armedAfterRender"] == "True", (
        "rendering a slider must arm the pump so drags are watchable at once"
    )
    assert facts["demandHolds"] == "True"
    assert facts["cleanedUp"] == "True"


def test_real_timer_end_to_end(run_async):
    facts = parse_transcript(run_async("TestRealTimerEndToEnd"))
    assert facts["armedAfterStart"] == "True"
    assert facts["unpinnedByDefault"] == "True", (
        "hover cursors must survive an armed pump by default"
    )
    assert facts["pinsOnRequest"] == "True"
    assert facts["unpinsOnRequest"] == "True"
    assert facts["statusNoManualPump"] == "done", (
        "the armed SetTimer alone must complete the op"
    )
    assert facts["autoDisarmed"] == "True", (
        "the pump must kill its own timer once work drains"
    )
