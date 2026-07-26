"""Live widget tests: progress, toggle, form controls, inputs, toasts, modal."""

from __future__ import annotations

import pytest

from conftest import _runner, parse_transcript


@pytest.fixture
def run_widgets(excel):
    return _runner(excel, "TestReDimWidgets")


def test_progress_bar(run_widgets):
    facts = parse_transcript(run_widgets("TestProgressBar"))
    assert facts["zeroHidesFill"] == "True"
    assert facts["halfVisible"] == "True"
    assert facts["halfWidthOk"] == "True"
    assert facts["fullWidthOk"] == "True"
    assert facts["clampedOk"] == "True"


def test_toggle(run_widgets):
    facts = parse_transcript(run_widgets("TestToggle"))
    assert facts["offFillMuted"] == "True"
    assert facts["stateOn"] == "True"
    assert facts["changeRan"] == "1"
    assert facts["onFillPrimary"] == "True"
    assert facts["knobMoved"] == "True"
    assert facts["checkedProp"] == "True"
    assert facts["knobClickTogglesOff"] == "True"


def test_select_box(run_widgets):
    facts = parse_transcript(run_widgets("TestSelectBox"))
    assert facts["faceText"] == "South"
    assert float(facts["faceFontSize"]) == 11.0
    assert facts["faceInkOnSurface"] == "True", (
        "face text must use surface ink, not the primary variant's white"
    )
    assert facts["caretExists"] == "True"
    assert facts["closedNoOptions"] == "True"
    assert facts["openOptions"] == "True"
    assert float(facts["optionFontSize"]) == 11.0, (
        "option list must use the theme font, unlike the native dropdown"
    )
    assert facts["optionText"] == "East"
    assert facts["pickedState"] == "East"
    assert facts["pickedFace"] == "East"
    assert facts["closedAfterPick"] == "True"
    assert facts["changeRan"] == "1"
    assert facts["toggleClosed"] == "True"


def test_text_input(run_widgets):
    facts = parse_transcript(run_widgets("TestTextInput"))
    assert facts["frameExists"] == "True"
    assert facts["apiWriteState"] == "Ada"
    assert facts["apiNoChangeProc"] == "True", (
        "API writes must not fire the change handler"
    )
    assert facts["editState"] == "Grace"
    assert facts["editChangeProc"] == "True"
    assert facts["inputReadback"] == "Grace"


def test_toast_lifecycle(run_widgets):
    facts = parse_transcript(run_widgets("TestToastLifecycle"))
    assert facts["shown"] == "True"
    assert facts["pendingWork"] == "True"
    assert facts["aliveBeforeTtl"] == "True"
    assert facts["removedAfterTtl"] == "True"
    assert facts["clickDismissed"] == "True"
    assert facts["lightInk"] == "True"
    assert facts["darkInk"] == "True", (
        "dark theme toast text must use the dark theme's surface ink"
    )


def test_toast_slots(run_widgets):
    facts = parse_transcript(run_widgets("TestToastSlots"))
    assert facts["entranceSlid"] == "True", (
        "the entrance slide must ease the toast up exactly its spawn offset"
    )
    assert facts["secondBelowFirst"] == "True"
    assert facts["survivorSlidUp"] == "True", (
        "dismissing a toast must slide the survivor up into the freed slot"
    )
    assert facts["thirdJoinsBelow"] == "True", (
        "a new toast must join below the compacted stack"
    )


def test_toast_tray(run_widgets):
    facts = parse_transcript(run_widgets("TestToastTray"))
    assert facts["onRail"] == "True"
    assert facts["modalIgnored"] == "True", (
        "modal chrome must not shift the toast rail"
    )
    assert facts["pinnedToAnchor"] == "True"
    assert facts["clampedIntoView"] == "True", (
        "a rail past the viewport edge must clamp into view"
    )
    assert facts["notAtRawRail"] == "True"


def test_drag_resilience(run_widgets):
    facts = parse_transcript(run_widgets("TestDragResilience"))
    assert facts["cardSwallowsClicks"] == "True", (
        "the modal card must carry the dispatcher so plain clicks cannot drag it"
    )
    assert facts["labelSwallowsClicks"] == "True"
    assert facts["cardSnappedBack"] == "True"
    assert facts["labelSnappedBack"] == "True"


def test_tick_box(run_widgets):
    facts = parse_transcript(run_widgets("TestTickBox"))
    assert facts["uncheckedSurface"] == "True"
    assert facts["glyphEmpty"] == "True"
    assert facts["captionText"] == "I agree"
    assert float(facts["captionSize"]) == 11.0
    assert facts["checkedState"] == "True"
    assert facts["checkedPrimary"] == "True"
    assert facts["glyphCheck"] == "True"
    assert facts["changeRan"] == "1"
    assert facts["captionToggles"] == "True", (
        "clicking the caption must toggle like the box"
    )


def test_radio_group(run_widgets):
    facts = parse_transcript(run_widgets("TestRadioGroup"))
    assert facts["rowsExist"] == "True"
    assert facts["dotOnSelected"] == "True"
    assert facts["dotOffOthers"] == "True"
    assert facts["captionText"] == "High"
    assert float(facts["captionSize"]) == 11.0
    assert facts["pickedState"] == "High"
    assert facts["dotMoved"] == "True"
    assert facts["changeRan"] == "1"
    assert facts["sameRowNoOp"] == "True"
    assert facts["rowOnePicked"] == "Low"


def test_stepper(run_widgets):
    facts = parse_transcript(run_widgets("TestStepper"))
    assert facts["faceValue"] == "4"
    assert facts["partsExist"] == "True"
    assert facts["plusValue"] == "5"
    assert facts["faceUpdated"] == "True"
    assert facts["clampedNoOp"] == "True", (
        "stepping past the maximum must not fire OnChange"
    )
    assert facts["minusValue"] == "4"
    assert facts["changeRan"] == "2"


def test_slide_bar(run_widgets):
    facts = parse_transcript(run_widgets("TestSlideBar"))
    assert facts["partsExist"] == "True"
    assert facts["fillFraction"] == "True"
    assert facts["thumbCentered"] == "True"
    assert facts["threeQuarterValue"] == "75"
    assert facts["changeRan"] == "1"
    assert facts["fillMoved"] == "True"
    assert facts["snappedValue"] == "50", "0.52 across 0..100 step 5 must snap to 50"
    assert facts["noExtraChange"] == "True"
    assert facts["minValue"] == "0"
    assert facts["fillHiddenAtMin"] == "True"
    assert facts["maxValue"] == "100"


def test_modal_confirm(run_widgets):
    facts = parse_transcript(run_widgets("TestModalConfirm"))
    assert facts["overlayShown"] == "True"
    assert facts["cardText"] == "True"
    assert facts["cancelHasBorder"] == "True", (
        "secondary buttons need an outline against surface backgrounds"
    )
    assert facts["overlayCoversOrigin"] == "True"
    assert facts["confirmRan"] == "1"
    assert facts["overlayHidden"] == "True"
    assert facts["cancelRan"] == "1"
    assert facts["confirmStillOne"] == "True"
    assert facts["overlayHiddenAgain"] == "True"
