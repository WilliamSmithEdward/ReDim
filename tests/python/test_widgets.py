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
    assert facts["repickQuiet"] == "True", (
        "re-picking the current item must close the list without firing"
        " OnChange"
    )
    assert facts["toggleClosed"] == "True"


def test_text_input(run_widgets):
    facts = parse_transcript(run_widgets("TestTextInput"))
    assert facts["frameExists"] == "True"
    assert facts["clickSelectsCell"] == "True", (
        "clicking a cell-backed input's frame must select its anchor cell"
    )
    assert facts["clickNoChange"] == "True", (
        "a click is not an edit and must not fire the change handler"
    )
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
    assert facts["fadesOut"] == "True", (
        "an expired toast must fade for a beat, not pop out of existence"
    )
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
    assert facts["closeRides"] == "True", (
        "the close button must move with the toast through its slide"
    )
    assert facts["secondBelowFirst"] == "True"
    assert facts["survivorSlidUp"] == "True", (
        "dismissing a toast must slide the survivor up into the freed slot"
    )
    assert facts["thirdJoinsBelow"] == "True", (
        "a new toast must join below the compacted stack"
    )
    assert facts["railSticksLeft"] == "True", (
        "the rail must not move while toasts live, even after scrolling"
    )
    assert facts["railSticksTop"] == "True", (
        "a newcomer must join the live column exactly one slot pitch below"
    )
    assert facts["glideEntryBelow"] == "True", (
        "a mid-glide newcomer must enter below the column as drawn, not"
        " on top of a survivor still easing through its slot"
    )
    assert facts["settledPitchA"] == "True"
    assert facts["settledPitchB"] == "True"
    assert facts["tickFaults"] == "0", (
        "the pump swallowed an error somewhere in this scenario's ticks"
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


def test_slide_drag(run_widgets):
    facts = parse_transcript(run_widgets("TestSlideDrag"))
    assert facts["dragWatchDemandsPump"] == "True", (
        "a slider on the active sheet must keep the pump armed for its press watch"
    )
    assert facts["dragging"] == "True"
    assert facts["pressValue"] == "30"
    assert facts["thumbAccent"] == "True", "the held thumb must show the accent"
    assert facts["liveValue"] == "90", "the value must track during the hold"
    assert facts["noChangeDuringHold"] == "True"
    assert facts["releasedFiredChange"] == "True"
    assert facts["thumbWhiteAgain"] == "True"
    assert facts["releaseClickSwallowed"] == "True", (
        "the OnAction click delivered at release must not re-set the value"
    )
    assert facts["noMoveNoChange"] == "True"


def test_slide_mapping(run_widgets):
    facts = parse_transcript(run_widgets("TestSlideMapping"))
    assert facts["scrollShiftMatches"] == "True", (
        "a fixed screen pixel must map to points shifted by exactly the scroll"
    )
    assert facts["zoomScales"] == "True", (
        "point spans must scale inversely with zoom"
    )
    assert facts["span100Sane"] == "True", (
        "300 px at 100 percent zoom must be roughly 225 points, not 300"
    )


def test_item_api(run_widgets):
    facts = parse_transcript(run_widgets("TestItemApi"))
    assert facts["countAfterInsert"] == "4"
    assert facts["insertedFirst"] == "Zeta"
    assert facts["selectionFollows"] == "True", (
        "an insert before the selection must keep the same item selected"
    )
    assert facts["selectionStillBeta"] == "True"
    assert facts["clearedToPlaceholder"] == "True", (
        "removing the selected item must clear to the placeholder"
    )
    assert facts["faceShowsPlaceholder"] == "True"
    assert facts["fromArray"] == "4"
    assert facts["fromRangeSkipsBlank"] == "3"
    assert facts["rangeSecond"] == "Green"
    assert facts["seqErrClean"] == "True", (
        "a ROneCOne item source must not trip the guarded array test or dirty Err"
    )
    assert facts["fromSequence"] == "3", (
        "a ROneCOne sequence must feed a picker, not fall through to no items"
    )
    assert facts["sequenceSecond"] == "Venus"
    assert facts["clearedNoOptions"] == "True"
    assert facts["radioShrunk"] == "2"
    assert facts["radioStaleGone"] == "True", (
        "shrinking a radio group must sweep its stale row parts"
    )
    assert facts["radioSelectionCleared"] == "True"


def test_combo_box(run_widgets):
    facts = parse_transcript(run_widgets("TestComboBox"))
    assert facts["frameAndCaret"] == "True"
    assert facts["faceSelectsCell"] == "True", (
        "clicking a cell-backed combo's face must select its anchor cell"
        " and open the list"
    )
    assert facts["openAll"] == "True"
    assert facts["optText"] == "Green"
    assert facts["pickedCell"] == "Green"
    assert facts["pickedState"] == "Green"
    assert facts["pickChangeRan"] == "1"
    assert facts["closedAfterPick"] == "True"
    assert facts["pickedIndex"] == "2"
    assert facts["reopenAll"] == "True", (
        "a picked value must reopen onto the whole list, not its one match"
    )
    assert facts["repickQuiet"] == "True", (
        "re-picking the item the cell holds must not fire OnChange"
    )
    assert facts["filteredCount"] == "True", (
        "the caret must open the list filtered by the typed text"
    )
    assert facts["filteredText"] == "Blue"
    assert facts["suggestOpened"] == "True", (
        "an Enter commit with partial text must auto-drop the suggestions"
    )
    assert facts["freeTextState"] == "gr"
    assert facts["exactClosed"] == "True"
    assert facts["exactIndex"] == "1"
    assert facts["noMatchClosed"] == "True"
    assert facts["noMatchState"] == "zzz"


def test_float_field(run_widgets):
    facts = parse_transcript(run_widgets("TestFloatField"))
    assert facts["focused"] == "True"
    assert facts["caretShown"] == "True", (
        "a focused float field must show the insertion bar"
    )
    assert facts["focusRing"] == "True"
    assert facts["inkOnSurface"] == "True", (
        "float field ink must be surface ink, not the primary variant's"
        " white-on-white"
    )
    assert facts["roundedFace"] == "True", (
        "field faces share the family's rounded profile"
    )
    assert facts["roundedCombo"] == "True"
    assert facts["typed"] == "True"
    assert facts["committed"] == "Hey"
    assert facts["changeRan"] == "1"
    assert facts["blurred"] == "True", "Enter must commit and release the keys"
    assert facts["plainText"] == "True"
    assert facts["reverted"] == "True", "Escape must revert and fire nothing"
    assert facts["comboOpenAll"] == "True"
    assert facts["liveFiltered"] == "True", (
        "the float combo must re-filter on every keystroke"
    )
    assert facts["narrowed"] == "True"
    assert facts["picked"] == "Green"
    assert facts["pickBlurred"] == "True"
    assert facts["pickClosed"] == "True"
    assert facts["reopenShowsAll"] == "True", (
        "refocusing after a pick must show every item, not only the pick"
    )
    assert facts["editFilters"] == "True", (
        "the first edit after reopening must filter the list again"
    )
    assert facts["escClosesList"] == "True", (
        "Esc on a combo must close its open list, keeping focus and text"
    )
    assert facts["escEscReverts"] == "True", (
        "a second Esc must revert the combo to the text focus found and"
        " leave"
    )
    assert facts["outsideCommit"] == "Heyo"
    assert facts["tabCommit"] == "Heyox", "Tab must commit exactly like Enter"
    assert facts["tabMovesOn"] == "True", (
        "Tab must move focus to the next field in Tab order"
    )
    assert facts["cellClickCommit"] == "Heyoxz", (
        "a selection change must commit the focused field - the cell-click"
        " press can be invisible to the pump's poll"
    )
    assert facts["cellClickBlurred"] == "True"


def test_transfer_list(run_widgets):
    facts = parse_transcript(run_widgets("TestTransferList"))
    assert facts["panels"] == "True"
    assert facts["buttons"] == "True"
    assert facts["leftRows"] == "True"
    assert facts["rightRows"] == "True"
    assert facts["headerCounts"] == "True", (
        "panel headers must show live item counts"
    )
    assert facts["rowSelected"] == "True"
    assert facts["selectedCheck"] == "True", (
        "a selected row must show a check, not only the accent fill"
    )
    assert facts["selectNoChange"] == "True", (
        "selecting a row must not fire OnChange; only transfers do"
    )
    assert facts["toggledOff"] == "True", (
        "clicking a selected row must toggle it back out of the set"
    )
    assert facts["multiSelected"] == "True", (
        "plain clicks accumulate a multi-selection - Excel reserves"
        " Ctrl+click on macro shapes, so toggling replaces the modifier"
    )
    assert facts["movedState"] == "Charlie, Delta, Alpha, Echo", (
        "moved items append to the chosen list in list order"
    )
    assert facts["movedCounts"] == "True"
    assert facts["changeRan"] == "1", "one multi-move fires OnChange once"
    assert facts["allRight"] == "True", (
        "move-all must empty the available panel and sweep its row shapes"
    )
    assert facts["allLeft"] == "True"
    assert facts["changeTotal"] == "3"
    assert facts["noSelNoOp"] == "True"
    assert facts["multiBack"] == "True"
    assert facts["changeFinal"] == "5"


def test_check_list(run_widgets):
    facts = parse_transcript(run_widgets("TestCheckList"))
    assert facts["parts"] == "True"
    assert facts["headerText"] == "Select all (1/3)", (
        "the select-all header must show the checked count"
    )
    assert facts["seeded"] == "True"
    assert facts["mixedDash"] == "True", (
        "a partially checked list must show the tri-state dash"
    )
    assert facts["captionToggled"] == "Alpha, Bravo", (
        "the caption is a full-row hit target and must toggle its row"
    )
    assert facts["changeRan"] == "1"
    assert facts["allChecked"] == "True", (
        "select-all from a mixed state must check everything"
    )
    assert facts["masterCheckGlyph"] == "True"
    assert facts["rowsChecked"] == "True", (
        "a row whose check changed must read fully checked: accent fill,"
        " check glyph, and glyph ink"
    )
    assert facts["noneChecked"] == "True", (
        "select-all from the all state must uncheck everything"
    )
    assert facts["rowsCleared"] == "True", (
        "an unchecked row must read empty: surface fill, no glyph, border"
    )
    assert facts["changeAfterMaster"] == "3", (
        "a bulk toggle fires OnChange once"
    )
    assert facts["mainShapeRow1"] == "True"
    assert facts["progSilent"] == "True", (
        "programmatic checks must not write state or fire OnChange"
    )
    assert facts["insertShift"] == "True", (
        "checks must follow their items through inserts"
    )
    assert facts["removeShift"] == "True", (
        "checks must follow their items through removals"
    )
    assert facts["captionSwap"] == "Xray,Zulu", (
        "new items at the same count must rewrite the captions"
    )
    assert facts["headerOff"] == "True"


def test_image(run_widgets):
    facts = parse_transcript(run_widgets("TestImage"))
    assert facts["picFill"] == "True", (
        "the image must render as a picture fill on the rounded shape"
    )
    assert facts["rounded"] == "True"
    assert facts["placeholder"] == "True", (
        "a missing source must render the themed placeholder"
    )
    assert facts["clickRan"] == "1"
    assert facts["boundSwap"] == "True", (
        "a state-bound source must swap the picture"
    )
    assert facts["embeddedKept"] == "True", (
        "the embedded picture must survive its source file's deletion"
    )


def test_multi_line_input(run_widgets):
    facts = parse_transcript(run_widgets("TestMultiLineInput"))
    assert facts["topAnchored"] == "True", (
        "a multi-line field must anchor its text to the top"
    )
    assert facts["enterStaysFocused"] == "True", (
        "Enter in a multi-line field must insert a newline, not commit"
    )
    assert facts["newlineInFace"] == "True"
    assert facts["ctrlEnterCommits"] == "True", (
        "Ctrl+Enter must commit the multi-line buffer with its newlines"
    )
    assert facts["changeRan"] == "1"
    assert facts["singleLineEnterCommits"] == "True", (
        "single-line fields keep the Enter-commits convention"
    )
    assert facts["tailWindow"] == "True", (
        "an overflowing multi-line field must window to the caret's tail"
        " lines with a leading ellipsis"
    )
    assert facts["fullCommit"] == "True", (
        "windowing is render-only; the committed value stays complete"
    )
    assert facts["lineWindow"] == "True", (
        "an overflowing single line must window to its rightmost chars"
    )
    assert facts["lineFullCommit"] == "True"


def test_caret_editing(run_widgets):
    facts = parse_transcript(run_widgets("TestCaretEditing"))
    assert facts["midInsert"] == "True", (
        "arrows must move the caret and characters must insert at it"
    )
    assert facts["bsAndDel"] == "True", (
        "Backspace deletes before the caret, Del deletes at it"
    )
    assert facts["homeJump"] == "True"
    assert facts["endJump"] == "True"
    assert facts["upClampsColumn"] == "True", (
        "vertical moves work in hard lines and keep the column"
    )
    assert facts["endOfLine"] == "True"
    assert facts["downClampsColumn"] == "True", (
        "a vertical move onto a shorter line clamps to its end"
    )
    assert facts["upScrolled"] == "True", (
        "the viewport must follow the caret upward with a trailing"
        " ellipsis for the lines below"
    )


def test_long_lists(run_widgets):
    facts = parse_transcript(run_widgets("TestLongLists"))
    assert facts["comboWindow"] == "True", (
        "an open combo must window to eight rows plus the bottom pager"
    )
    assert facts["moreText"] == "▼ 4 more"
    assert facts["scrolledRow1"] == "Item02", (
        "walking the highlight past the window edge must scroll it"
    )
    assert facts["highlightLast"] == "True"
    assert facts["enterTakes"] == "Item09", (
        "Enter must take the highlighted match"
    )
    assert facts["clickMaps"] == "Item02", (
        "clicked window rows must map through the scroll offset"
    )
    assert facts["comboPagedRow1"] == "Item05", (
        "the bottom pager must page the window, clamped to the tail"
    )
    assert facts["optuText"] == "▲ 4 more"
    assert facts["optdGoneAtEnd"] == "True", (
        "the bottom pager disappears when nothing lies below"
    )
    assert facts["downStartsInWindow"] == "True", (
        "a Down after mouse paging starts the highlight in the window"
    )
    assert facts["pagedBack"] == "True"
    assert facts["reopenRow1"] == "Item03", (
        "a combo must reopen onto the whole list scrolled to its pick"
    )
    assert facts["reopenWindowed"] == "True"
    assert facts["downFromPick"] == "True", (
        "Down on a reopened list must walk on from the pick"
    )
    assert facts["typingFilters"] == "True"
    assert facts["keyPickFires"] == "1", (
        "an Enter pick that changes the value must fire OnChange once"
    )
    assert facts["keyPickState"] == "Item01"
    assert facts["keyPickNoRefire"] == "1", (
        "the blur after a keyboard pick must not fire OnChange again"
    )
    assert facts["keyRepickQuiet"] == "0", (
        "an Enter re-pick of the current value must not fire OnChange"
    )
    assert facts["mouseRepickQuiet"] == "0", (
        "a mouse re-pick of the current value must not fire OnChange"
    )
    assert facts["mousePickFires"] == "1"
    assert facts["mousePickState"] == "Item03"
    assert facts["comboListRows"] == "True", "ListRows must size the combo window"
    assert facts["selectWindow"] == "True", (
        "an open SelectBox must window to eight rows behind pagers"
    )
    assert facts["selectShowsPick"] == "Item10", (
        "opening a SelectBox must scroll its selection into view"
    )
    assert facts["selectPagedBack"] == "True"
    assert facts["selectRowMaps"] == "Item02", (
        "clicked SelectBox rows must map through the scroll offset"
    )
    assert facts["selectListRows"] == "True", (
        "ListRows must size the SelectBox window"
    )
    assert facts["listRowsKindGuard"] == "True"
    assert facts["listRowsMinimum"] == "True"
    assert facts["poolRows"] == "True"
    assert facts["pagerTarget"] == "True", (
        "transfer paging arrows must be at least 18 points square, clear"
        " of the rows"
    )
    assert facts["pagedRow1"] == "Item05", (
        "the panel scroll button must page the window"
    )
    assert facts["offsetMovedState"] == "Item06", (
        "row selection must map through the panel scroll offset"
    )
    assert facts["scrollerGoneWhenFits"] == "True"


def test_chrome_claim(run_widgets):
    facts = parse_transcript(run_widgets("TestChromeClaim"))
    assert facts["closedClaimsNothing"] == "True"
    assert facts["openClaims"] == "True", (
        "an open drop list must claim the points it covers so watches"
        " cannot hit-test through it"
    )
    assert facts["besideNotClaimed"] == "True"
    assert facts["exceptSelf"] == "True"
    assert facts["closedAgain"] == "True"
    assert facts["selectClaimsWindow"] == "True"
    assert facts["selectBelowWindowFree"] == "True", (
        "an open SelectBox claims its windowed depth, not every item"
    )
    assert facts["overlayClaims"] == "True", (
        "a modal overlay must claim everything it covers"
    )
    assert facts["overlayReleased"] == "True"


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


def test_list_dismiss(run_widgets):
    facts = parse_transcript(run_widgets("TestListDismiss"))
    assert facts["oneListPerApp"] == "True", (
        "opening a list must close the app's other open list"
    )
    assert facts["clickElsewhereCloses"] == "True", (
        "a click on another control must close an open list"
    )
    assert facts["pressOnRowsKeeps"] == "True", (
        "a press on the list's own rows must leave it open"
    )
    assert facts["pressOffCloses"] == "True", (
        "a press off the face and rows must close the list"
    )
    assert facts["selectionHeldKeeps"] == "True"
    assert facts["selectionMoveCloses"] == "True", (
        "moving the grid selection must close an open list"
    )


def test_list_conventions(run_widgets):
    facts = parse_transcript(run_widgets("TestListConventions"))
    assert facts["currentChecked"] == "True", (
        "only the current item's row carries the check"
    )
    assert facts["rowItem"] == "Green"
    assert facts["noItemsRow"] == "No items"
    assert facts["emptyRowInert"] == "True", (
        "clicking the empty row must leave the list as it is"
    )
    assert facts["noMatchesRow"] == "No matches"
    assert facts["emptyRowGone"] == "True"
    assert facts["opensUpward"] == "True", (
        "a list without room below the window must open above its face"
    )
    assert facts["firstRowOnTop"] == "True", (
        "an upward list keeps its items in order, first on top"
    )
    assert facts["upwardPick"] == "True"


def test_motion_and_blink(run_widgets):
    facts = parse_transcript(run_widgets("TestMotionAndBlink"))
    assert facts["motionReduced"] == "True"
    assert facts["noEntranceSlide"] == "True", (
        "with motion reduced a toast must appear in its slot"
    )
    assert facts["leavesAtOnce"] == "True", (
        "with motion reduced a dismissed toast must leave without a fade"
    )
    assert facts["survivorMovesAtOnce"] == "True", (
        "with motion reduced a survivor must move to its new slot at once"
    )
    assert facts["overrideFull"] == "True"
    assert facts["blinkAtSystemRate"] == "True", (
        "the caret must blink at the Windows GetCaretBlinkTime interval"
    )


def test_accessibility(run_widgets):
    facts = parse_transcript(run_widgets("TestAccessibility"))
    assert facts["buttonAlt"] == "Save, button"
    assert facts["disabledAlt"] == "Save, button, unavailable"
    assert facts["tickAlt"] == "Agree, checkbox, checked"
    assert facts["toggleAlt"] == "Switch, off"
    assert facts["selectAlt"] == "Drop-down, South"
    assert facts["progressAlt"] == "Progress, 40 percent"
    assert facts["overrideAlt"] == "Dark mode", (
        "AltText must replace the generated description"
    )
    assert facts["blackOnWhite"] == "21.00"
    assert facts["sameColor"] == "1.00"
    assert facts["highContrastFails"] == "0", (
        "every pairing in the high-contrast theme must pass WCAG AA"
    )
    assert facts["reportLines"] == "9"
    assert facts["knobOnAccent"] == "True"


def test_toast_conventions(run_widgets):
    facts = parse_transcript(run_widgets("TestToastConventions"))
    assert facts["closeButton"] == "True", "every toast needs a visible close"
    assert facts["plainText"] == "Saved.", (
        "a toast with no tone must carry no icon"
    )
    assert facts["shortTtl"] == "True", "short messages stay four seconds"
    assert facts["closeDismisses"] == "True"
    assert facts["longTtl"] == "True", (
        "a 100-character message stays three seconds plus 60 ms a character"
    )
    assert facts["warningIcon"] == "True"
    assert facts["toneColors"] == "True", (
        "a tone colors the icon and the edge; the close glyph stays muted"
    )
    assert facts["actionText"] == "Undo"
    assert facts["actionLongerTtl"] == "True", (
        "an action adds four seconds to the toast's time"
    )
    assert facts["actionRan"] == "True"
    assert facts["actionDismisses"] == "True"


def test_keyboard_focus(run_widgets):
    facts = parse_transcript(run_widgets("TestKeyboardFocus"))
    assert facts["firstFocused"] == "True", "TabIndex 1 must come first"
    assert facts["ringDrawn"] == "True", "a focused button must wear a ring"
    assert facts["tabToField"] == "True", (
        "Tab must move on in creation order and take the ring along"
    )
    assert facts["tabCommits"] == "Bo"
    assert facts["onButton"] == "True"
    assert facts["spaceClicks"] == "True", "Space must click a focused button"
    assert facts["spaceToggles"] == "True"
    assert facts["spaceChecks"] == "True"
    assert facts["tabWraps"] == "True", "Tab must wrap, skipping TabIndex -1"
    assert facts["backTab"] == "True"
    assert facts["focusApi"] == "True"
    assert facts["escLeaves"] == "True"
    assert facts["clickNoFocus"] == "True", (
        "a mouse click must not give a button keyboard focus"
    )
    assert facts["enterDefault"] == "True", (
        "Enter in a field must commit and click the default button"
    )
    assert facts["keyUnderlined"] == "True"
    assert facts["accessKeyClicks"] == "True"
    assert facts["selectionEndsFocus"] == "True"
    assert facts["clickEndsFocus"] == "True"
    assert facts["scrolledIntoView"] == "True", (
        "focus below the window must scroll the control into view"
    )


def test_control_keys(run_widgets):
    facts = parse_transcript(run_widgets("TestControlKeys"))
    assert facts["radioDown"] == "Medium"
    assert facts["radioWraps"] == "Large", "Up from the first row must wrap"
    assert facts["radioFires"] == "3"
    assert facts["stepUp"] == "6"
    assert facts["stepPage"] == "16"
    assert facts["stepEnd"] == "20"
    assert facts["stepHome"] == "0"
    assert facts["slider"] == "45", "Right steps 5, Page Down a tenth (10)"
    assert facts["sliderEnd"] == "100"
    assert facts["cursorDrawn"] == "True"
    assert facts["checkKeys"] == "True"
    assert facts["headerKey"] == "True", "Space on the header checks all"
    assert facts["enterMoves"] == "True", (
        "Enter must move the cursor's row when nothing is selected"
    )
    assert facts["movesBack"] == "True"
    assert facts["selectArrow"] == "2", "Down on a closed select selects next"
    assert facts["typeAhead"] == "5", "n from South must land on Northwest"
    assert facts["spaceOpens"] == "True", (
        "Space must open the list with the current item highlighted"
    )
    assert facts["openPick"] == "True"
    assert facts["escCloses"] == "True"
    assert facts["escLeaves"] == "True"


def test_modal_keys(run_widgets):
    facts = parse_transcript(run_widgets("TestModalKeys"))
    assert facts["modalTakesFocus"] == "True"
    assert facts["tabToCancel"] == "True"
    assert facts["trapped"] == "True", "Tab must stay among the modal's buttons"
    assert facts["enterConfirms"] == "True"
    assert facts["focusReturns"] == "True", (
        "closing the modal must return focus to the control that had it"
    )
    assert facts["escCancels"] == "True"


def test_clearable(run_widgets):
    facts = parse_transcript(run_widgets("TestClearable"))
    assert facts["hiddenWhenEmpty"] == "True"
    assert facts["shownWithText"] == "True"
    assert facts["clearEmpties"] == "True", (
        "the clear button must empty the field and keep focus"
    )
    assert facts["comboX"] == "True"
    assert facts["comboCleared"] == "True", (
        "clearing a combo must open its whole list"
    )


def test_text_editing(run_widgets):
    facts = parse_transcript(run_widgets("TestTextEditing"))
    assert facts["selectionPainted"] == "True", (
        "a Shift selection must show as an accent highlight"
    )
    assert facts["cut"] == "hello "
    assert facts["undo"] == "hello world"
    assert facts["redo"] == "hello "
    assert facts["pasted"] == "world", "Ctrl+V must replace the selection"
    assert facts["pasteUndone"] == "ab"
    assert facts["typingOneStep"] == "True", "a typing run undoes as one step"
    assert facts["wordLeft"] == "one two _three"
    assert facts["wordBackspace"] == "one two ", (
        "Ctrl+Backspace deletes a word; the underscore is part of it"
    )
    assert facts["shiftWordDelete"] == "two ", (
        "Ctrl+Shift+Right selects to the next word start"
    )
    assert facts["symbols"] == "@'\"(~"


def test_field_clicks(run_widgets):
    facts = parse_transcript(run_widgets("TestFieldClicks"))
    assert facts["clickPlacesCaret"] == "alpha Xbeta gamma", (
        "the focusing click must put the caret under the pointer"
    )
    assert facts["clickMovesCaret"] == "alYpha Xbeta gamma"
    assert facts["doubleClickWord"] == "alYpha Xbeta Z", (
        "a double click must select the word under the pointer"
    )


def test_field_rules(run_widgets):
    facts = parse_transcript(run_widgets("TestFieldRules"))
    assert facts["placeholderShown"] == "True", (
        "an empty field must show its placeholder in muted ink"
    )
    assert facts["placeholderFocused"] == "True"
    assert facts["placeholderGone"] == "True"
    assert facts["numeric"] == "True", (
        "Numeric must keep digits, one separator, and a leading minus"
    )
    assert facts["maxLength"] == "abcde"
    assert facts["counter"] == "5/5"
    assert facts["invalid"] == "Needs an @"
    assert facts["messageShown"] == "True"
    assert facts["dangerBorder"] == "True"
    assert facts["liveRecheck"] == "True", (
        "an invalid field must recheck on every edit and clear once valid"
    )
    assert facts["inputNow"] == "2:xy"
    assert facts["debounceWaits"] == "True"
    assert facts["debounceFires"] == "3:xyz"


def test_auto_grow(run_widgets):
    facts = parse_transcript(run_widgets("TestAutoGrow"))
    assert facts["startsAtGiven"] == "True"
    assert facts["firstRenderGrown"] == "True", (
        "a first render must size an AutoGrow field to its text"
    )
    assert facts["grewTwoLines"] == "True"
    assert facts["belowFollows"] == "True", (
        "a component placed Below a growing field must move with it"
    )
    assert facts["capped"] == "True"
    assert facts["faceScrolls"] == "True"
    assert facts["shrinksBack"] == "True"
    assert facts["belowReturns"] == "True"
    assert facts["wrapGrows"] == "True"


def test_combo_assists(run_widgets):
    facts = parse_transcript(run_widgets("TestComboAssists"))
    assert facts["boldMatch"] == "B,B", (
        "drop rows must bold what the typed text matched"
    )
    assert facts["ghostShown"] == "True", (
        "the first item the text begins must show its rest after the caret"
    )
    assert facts["ghostMuted"] == "True"
    assert facts["boldGrows"] == "Ba", (
        "a row that keeps its item must bold a longer match"
    )
    assert facts["boldShrinks"] == "B"
    assert facts["boldCleared"] == "True"
    assert facts["ghostFollows"] == "True:Bl"
    assert facts["rightTakes"] == "Blueberry"
    assert facts["undoGivesBack"] == "bl"
    assert facts["tabTakes"] == "Cherry:True"
    assert facts["restrictReverts"] == "True", (
        "a restricted combo must revert a commit that names no item"
    )
    assert facts["restrictCompletes"] == "Blueberry"
    assert facts["restrictSpelling"] == "Apple"
    assert facts["restrictEmpty"] == "True"


def test_pointer_basics(run_widgets):
    facts = parse_transcript(run_widgets("TestPointerBasics"))
    assert facts["pressDrags"] == "True:50", (
        "a held button on the track must start a drag through the pump"
    )
    assert facts["bubbleShown"] == "50"
    assert facts["bubbleAbove"] == "True"
    assert facts["bubbleFollows"] == "80"
    assert facts["releaseCommits"] == "1:80"
    assert facts["bubbleGone"] == "True"
    assert facts["focusBubble"] == "90"
    assert facts["blurHides"] == "True"
    assert facts["toastHeld"] == "True", (
        "a toast under the pointer must stop counting down"
    )
    assert facts["toastResumes"] == "True"


def test_pointer_effects(run_widgets):
    facts = parse_transcript(run_widgets("TestPointerEffects"))
    assert facts["buttonHover"] == "True", "a hovered button must tint"
    assert facts["buttonPressed"] == "True"
    assert facts["releaseHover"] == "True"
    assert facts["crossingNotPressed"] == "True", (
        "a press that went down elsewhere must not press what it crosses"
    )
    assert facts["stepperPart"] == "True"
    assert facts["buttonCleared"] == "True"
    assert facts["checkRow"] == "True"
    assert facts["transferRow"] == "True"
    assert facts["transferButton"] == "True"
    assert facts["rowCleared"] == "True"
    assert facts["listFollows"] == "True", (
        "an open list's highlight must follow a moving pointer"
    )
    assert facts["effectsOff"] == "True"
    assert facts["rearmed"] == "True", (
        "a sheet coming back to the front must re-arm the pump"
    )


def test_tooltips(run_widgets):
    facts = parse_transcript(run_widgets("TestTooltips"))
    assert facts["altTip"] == "Save, button, Saves the form"
    assert facts["altReason"] == "Send, button, unavailable, Fill in the address first"
    assert facts["waits"] == "True", "a tooltip must wait for the pointer to rest"
    assert facts["tipShows"] == "Saves the form"
    assert facts["belowPointer"] == "True"
    assert facts["pressHides"] == "True", (
        "a press must put the tooltip away until the pointer leaves"
    )
    assert facts["reasonShows"] == "Fill in the address first"
    assert facts["leaveHides"] == "True"


def test_hold_repeat(run_widgets):
    facts = parse_transcript(run_widgets("TestHoldRepeat"))
    assert facts["noStepAtPress"] == "5"
    assert facts["firstRepeat"] == "6", (
        "a held stepper button must step after the keyboard repeat delay"
    )
    assert facts["keepsRepeating"] == "7"
    assert facts["stateLive"] == "7"
    assert facts["noChangeYet"] == "0"
    assert facts["releaseChange"] == "1", "the release must fire OnChange once"
    assert facts["releaseClickSwallowed"] == "7"
    assert facts["tapSteps"] == "8:2"
    assert facts["arrowPages"] == "Item 11", (
        "a held paging arrow must page again and again"
    )


def test_transfer_gestures(run_widgets):
    facts = parse_transcript(run_widgets("TestTransferGestures"))
    assert facts["doubleClickMoves"] == "B:C:1", (
        "a double click must move the row across and fire OnChange once"
    )
    assert facts["dropOutlined"] == "True"
    assert facts["dragMoves"] == "B, A:2"
    assert facts["outlineCleared"] == "True"
    assert facts["rangeSelected"] == "True", (
        "a press dragged along a panel must select the rows it covers"
    )
    assert facts["rangeMoves"] == "B, A, C, D, E"


def test_item_values(run_widgets):
    facts = parse_transcript(run_widgets("TestItemValues"))
    assert facts["selectValue"] == "20:Integer", (
        "a pick must write the item's value, with its own type"
    )
    assert facts["itemValueAt"] == "30:"
    assert facts["radioValue"] == "P"
    assert facts["comboPick"] == "#f00"
    assert facts["comboFreeText"] == "Blue"
    assert facts["checkValues"] == "1, 3"
    assert facts["transferValues"] == "y1:y1"
    assert facts["replacedDropsValues"] == "One"
    assert facts["chosenKeepValue"] == "y1"


def test_select_groups(run_widgets):
    facts = parse_transcript(run_widgets("TestSelectGroups"))
    assert facts["headerLook"] == "True", "a group header must read bold and muted"
    assert facts["disabledLook"] == "True"
    assert facts["clicksIgnored"] == "True", (
        "clicks on a header or a disabled item must do nothing"
    )
    assert facts["picked"] == "Apple"
    assert facts["downSkips"] == "Carrot", "Down must pass over disabled items and headers"
    assert facts["upSkips"] == "Apple"
    assert facts["homeSkipsHeader"] == "2"
    assert facts["typeAheadSkips"] == "Apple"
    assert facts["headerNotValue"] == "0"
    assert facts["marksShift"] == "True"


def test_transfer_reorder(run_widgets):
    facts = parse_transcript(run_widgets("TestTransferReorder"))
    assert facts["arrowsDrawn"] == "True"
    assert facts["upOnce"] == "A, C, D, B", "the selected rows must move up as a block"
    assert facts["stopsAtTop"] == "C, D, A, B:2", (
        "a move at the top must change nothing and fire nothing"
    )
    assert facts["selectionFollows"] == "True"
    assert facts["downMoves"] == "A, C, D, B"
    assert facts["altUpMovesCursorRow"] == "A, C, B, D"
    assert facts["cursorFollows"] == "A, B, C, D"


def test_list_filter(run_widgets):
    facts = parse_transcript(run_widgets("TestListFilter"))
    assert facts["headerShowsFilter"] == 'Available "ap" (2 of 5)', (
        "typing while a transfer list has the keys must filter its panel"
    )
    assert facts["rowsFiltered"] == "Apple,Apricot:False"
    assert facts["keysOnShownRows"] == "Apricot"
    assert facts["backspaceWidens"] == 'Available "a" (2 of 4)'
    assert facts["moveAllShown"] == "Apricot, Apple, Banana", (
        "move-all under a filter must move only the rows that show"
    )
    assert facts["escClears"] == "Available (2):True"
    assert facts["checkHeader"] == 'Select all "gr" (0/2)'
    assert facts["checkRowsHidden"] == "True"
    assert facts["selectAllShown"] == "Green, Gray"
    assert facts["allBack"] == "True"


def test_adornments(run_widgets):
    facts = parse_transcript(run_widgets("TestAdornments"))
    assert facts["captionAbove"] == "True", (
        "a caption must sit above the field with a danger-colored required mark"
    )
    assert facts["hintBelow"] == "True"
    assert facts["selectCaption"] == "True"
    assert facts["requiredShown"] == "True", (
        "an empty commit of a required field must show its message"
    )
    assert facts["requiredClears"] == "True"
    assert facts["errorTextShown"] == "True"
    assert facts["errorTextCleared"] == "True"
    assert facts["skeletonBars"] == "True"
    assert facts["skeletonPulses"] == "True"


def test_tabs(run_widgets):
    facts = parse_transcript(run_widgets("TestTabs"))
    assert facts["firstPanel"] == "True", (
        "the first tab must show its panel and hide the others"
    )
    assert facts["firstBold"] == "True"
    assert facts["barUnderFirst"] == "True"
    assert facts["clickSwitches"] == "True"
    assert facts["writes"] == "Advanced"
    assert facts["fires"] == "1"
    assert facts["hiddenStays"] == "True", (
        "a control set Visible False stays hidden on its own tab"
    )
    assert facts["altText"] == "True"
    assert facts["visibleOnTab"] == "True"
    assert facts["keyRight"] == "True"
    assert facts["visibleWaited"] == "True", (
        "Visible set while the tab hid the control must apply when the tab shows"
    )
    assert facts["keyWraps"] == "1"
    assert facts["keyEnd"] == "3"
    assert facts["hoverTints"] == "True"
    assert facts["hidingCommits"] == "True", (
        "a focused field on a panel that hides must commit and let focus go"
    )
    assert facts["tabsHidden"] == "True"
    assert facts["tabsBack"] == "True"
    assert facts["removedShowsAll"] == "True"


def test_date_picker(run_widgets):
    facts = parse_transcript(run_widgets("TestDatePicker"))
    assert facts["placeholder"] == "True", "the placeholder must read muted"
    assert facts["faceFormat"] == "2026-09-22"
    assert facts["picked"] == "2026-09-22"
    assert facts["programSilent"] == "True", (
        "PickDate must write nothing and fire nothing"
    )
    assert facts["opens"] == "True", "a click must open the calendar on the date's month"
    assert facts["dateFilled"] == "True"
    assert facts["outOfRangeKeeps"] == "True", "a day outside DateRange must not pick"
    assert facts["nextMonth"] == "True"
    assert facts["dayPicks"] == "True", (
        "a day must write a Date, fire OnChange once, and close the calendar"
    )
    assert facts["written"] == "2026-10-15"
    assert facts["faceShows"] == "2026-10-15"
    assert facts["keyOpens"] == "True"
    assert facts["keysPick"] == "2026-11-23/2", (
        "Right, Down, and Page Down walk a day, a week, and a month; Enter picks"
    )
    assert facts["escCloses"] == "True"
    assert facts["pressOnKeeps"] == "True"
    assert facts["hoverTints"] == "True"
    assert facts["pressOffCloses"] == "True"
