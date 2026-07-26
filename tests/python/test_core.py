"""Live core tests: mount, render, diffing, dispatch, guards, state, theme."""

from __future__ import annotations

from conftest import parse_transcript


def test_mount_and_render(run_core):
    facts = parse_transcript(run_core("TestMountAndRender"))
    assert facts["shapes"] == "3"
    assert facts["buttonText"] == "Run"
    assert facts["fillIsPrimary"] == "True"
    assert facts["onAction"].endswith("!RdxDispatch")
    assert facts["labelText"] == "Ready"
    assert facts["geometry"] == "True"


def test_idempotent_remount(run_core):
    facts = parse_transcript(run_core("TestIdempotentRemount"))
    assert facts["first"] == "2"
    assert facts["second"] == "2", "same-session remount must not duplicate shapes"
    assert facts["third"] == "2", "post-state-loss rebuild must adopt shapes"
    assert facts["rebuiltText"] == "Go rebuilt"


def test_dispatch_and_sender(run_core):
    facts = parse_transcript(run_core("TestDispatchAndSender"))
    assert facts["clicks"] == "1"
    assert facts["senderId"] == "btnGo"
    assert facts["senderApp"] == "core3"
    assert facts["senderCleared"] == "True"
    assert facts["missingIgnored"] == "True"


def test_click_guards(run_core):
    facts = parse_transcript(run_core("TestClickGuards"))
    assert facts["afterDoubleClick"] == "1", "debounce must swallow the second click"
    assert facts["afterDebounceWait"] == "2"
    assert facts["afterDisabledClick"] == "2", "disabled button must not fire"
    assert facts["disabledFillMuted"] == "True"
    assert facts["afterReEnabled"] == "3"


def test_state_bindings(run_core):
    facts = parse_transcript(run_core("TestStateBindings"))
    assert facts["boundText"] == "Status: starting"
    assert facts["disabledClicks"] == "0"
    assert facts["updatedText"] == "Status: ready"
    assert facts["hintHidden"] == "True"
    assert facts["enabledClicks"] == "1"
    assert facts["stateReadback"] == "ready"
    assert facts["invertedIdleEnabled"] == "True"
    assert facts["invertedBusyDisabled"] == "True"


def test_batch_and_theme(run_core):
    facts = parse_transcript(run_core("TestBatchAndTheme"))
    assert facts["batchedText"] == "batched"
    assert facts["darkFill"] == "True"
    assert facts["darkIsDark"] == "True"


def test_state_handlers(run_core):
    facts = parse_transcript(run_core("TestStateHandlers"))
    assert facts["handlersRanOnSet"] == "2", (
        "both registered handlers must fire on SetState"
    )
    assert facts["handlersRanAgain"] == "4"
    assert facts["unwatchedIgnored"] == "4"


def test_unmount(run_core):
    facts = parse_transcript(run_core("TestUnmount"))
    assert facts["before"] == "2"
    assert facts["after"] == "0"
    assert facts["forgotten"] == "True"


def test_persistence(run_core):
    facts = parse_transcript(run_core("TestPersistence"))
    assert facts["beforeLoss"] == "dark"
    assert facts["modeKept"] == "dark", "defaults must not clobber persisted values"
    assert facts["runsKept"] == "7"
    assert facts["runsType"] == "Long"
    assert facts["ratioKept"] == "True"
    assert facts["flagKept"] == "True"
    assert facts["flagType"] == "Boolean"
    assert facts["clearedGone"] == "True"


def test_relative_layout(run_core):
    facts = parse_transcript(run_core("TestRelativeLayout"))
    assert facts["underLeftAligned"] == "True"
    assert facts["underBelow"] == "True"
    assert facts["asideTopAligned"] == "True"
    assert facts["asideRight"] == "True"
    assert facts["badRefErr"] == "True"


def test_orphan_pruning(run_core):
    facts = parse_transcript(run_core("TestOrphanPruning"))
    assert facts["beforeCount"] == "2"
    assert facts["afterCount"] == "2", "renamed component must not leave a third shape"
    assert facts["oldGone"] == "True"
    assert facts["newExists"] == "True"


def test_hotkey_lifecycle_and_version(run_core):
    facts = parse_transcript(run_core("TestHotKeyLifecycle"))
    assert facts["procCallable"] == "True"
    assert facts["unmountClean"] == "True"
    assert facts["version"] == "0.2.4"
