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


def test_theme_swap_repaint(run_core):
    facts = parse_transcript(run_core("TestThemeSwapRepaint"))
    assert facts["lightCanvas"] == "True"
    assert facts["darkCanvas"] == "True", (
        "SetTheme must repaint a painted canvas background"
    )
    assert facts["darkCard"] == "True", "cards must retint on a theme swap"
    assert facts["darkTrack"] == "True", (
        "the progress track must retint on a theme swap"
    )
    assert facts["swapChanged"] == "True"


def test_state_handlers(run_core):
    facts = parse_transcript(run_core("TestStateHandlers"))
    assert facts["handlersRanOnSet"] == "11", (
        "both listeners must fire on SetState, a repeated registration only once"
    )
    assert facts["handlersRanAgain"] == "22"
    assert facts["unwatchedIgnored"] == "22"
    assert facts["secondKeyHeard"] == "32", "a listener given two keys must hear the second"


def test_unmount(run_core):
    facts = parse_transcript(run_core("TestUnmount"))
    assert facts["before"] == "2"
    assert facts["after"] == "0"
    assert facts["forgotten"] == "True"


def test_state_session_scope(run_core):
    facts = parse_transcript(run_core("TestStateSessionScope"))
    assert facts["defaultNoClobber"] == "dark", (
        "SetStateDefault must not overwrite a value already in play"
    )
    assert facts["freshStoreEmpty"] == "True", (
        "state is session-scoped: a rebuilt app starts with an empty store"
    )
    assert facts["defaultSeeds"] == "light"


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


def test_navigation(run_core):
    facts = parse_transcript(run_core("TestNavigation"))
    assert facts["activeA"] == "True"
    assert facts["aVisible"] == "True"
    assert facts["bHidden"] == "True", "inactive windows must be very-hidden"
    assert facts["cHidden"] == "True"
    assert facts["navBarTabs"] == "True", "the bar must show one tab per window"
    assert facts["activeTabPrimary"] == "True"
    assert facts["tabTitleText"] == "Alpha", "tabs must show WindowTitle"
    assert facts["activeB"] == "True", "NavigatesTo must switch windows"
    assert facts["aNowHidden"] == "True"
    assert facts["log"] == "+A-A+B", "lifecycle hooks must fire in order"
    assert facts["tabSwapped"] == "True", (
        "the highlight must follow navigation onto every bar"
    )
    assert facts["activeC"] == "True"
    assert facts["backToB"] == "True"
    assert facts["activeAfterBack"] == "navb"
    assert facts["backToA"] == "True"
    assert facts["backEmpty"] == "True", "an empty back stack must report False"
    assert facts["plainRefused"] == "True"
    assert facts["goneRefused"] == "True"
    assert facts["staleTabPruned"] == "True"
    assert facts["liveTabsRemain"] == "True"
    assert facts["shutdownClearsActive"] == "True", (
        "Shutdown must forget the active window with the apps"
    )
    assert facts["rebuildAfterShutdown"] == "True", (
        "a NavBar built after Shutdown must not tab to unmounted windows"
    )
    assert facts["backStackCleared"] == "True", (
        "Shutdown must empty the back stack"
    )


def test_protect_surface(run_core):
    facts = parse_transcript(run_core("TestProtectSurface"))
    assert facts["protected"] == "True"
    assert facts["inputUnlocked"] == "True", (
        "TextInput cells must stay editable under protection"
    )
    assert facts["otherCellsLocked"] == "True"
    assert facts["selectionLocked"] == "True", (
        "locked canvas cells must be unselectable on a protected surface"
    )
    assert facts["renderWorks"] == "True", (
        "UserInterfaceOnly must leave framework writes free"
    )
    assert facts["protectedListOpens"] == "True", (
        "a drop list must open on a protected surface and leave it protected as it was"
    )
    assert facts["protectedCalendarOpens"] == "True", (
        "a calendar must open on a protected surface and leave it protected as it was"
    )
    assert facts["dispatchWorks"] == "True"
    assert facts["toastCreates"] == "True"
    assert facts["unprotects"] == "True"
    assert facts["selectionRestored"] == "True"
    assert facts["optOutSelectable"] == "True", (
        "allowCellSelection:=True must keep the grid selectable"
    )
    assert facts["unmountUnprotects"] == "True"


def test_writes_to_follows(run_core):
    facts = parse_transcript(run_core("TestWritesToFollows"))
    assert facts["seeded"] == "True/0", (
        "a key with no value takes the control's on render, firing no listener"
    )
    assert facts["follows"] == "True/1", (
        "SetState must move the control; the listener runs, OnChange does not"
    )
    assert facts["valueMaps"] == "2", "an item's value must map back to its item"
    assert facts["textMaps"] == "1", "an item's text must map back to its item"
    assert facts["clamps"] == "50"
    assert facts["fieldFollows"] == "Ada"
    assert facts["typingKept"] == "Adax", "a focused field must keep what is being typed"
    assert facts["renderKeepsState"] == "True", "Render must show the value in play"


def test_writes_to_edges(run_core):
    facts = parse_transcript(run_core("TestWritesToEdges"))
    assert facts["seedShown"] == "Volume 40", (
        "a label drawn before the control that seeds its key must show the seed"
    )
    assert facts["listsSeed"] == "Alpha, Charlie/Dee", (
        "a check list and a transfer list must seed their keys on the first render"
    )
    assert facts["positions"] == "2/1/0", "ItemPosition and ChosenPosition ignore case"
    assert facts["sharedValueKept"] == "2/1", (
        "a pick among items sharing a value must stay on the item picked"
    )
    assert facts["zeroNames"] == "3"
    assert facts["emptyClears"] == "0", "Empty must clear the pick, not name an item valued 0"
    assert facts["nullClears"] == "0"
    assert facts["drawsAfterNull"] == "30/Volume 30", "a Null key must not stop later draws"
    assert facts["clampKept"] == "100", "a same-key BindValue must not undo the clamp"
    assert facts["dateFromText"] == "True", "a numeric string must name a date"
    assert facts["hugeIgnored"] == "True", "a number past the date range must be ignored"
    assert facts["listenerSees"] == "True", "a listener must read the value the control shows"
    assert facts["writeOverrules"] == "True", "writing the key must overrule a setter"
    assert facts["requiredClears"] == "True", "state filling a pick must clear Required"


def test_focus_rails(run_core):
    facts = parse_transcript(run_core("TestFocusRails"))
    assert facts["captured"] == "True", "a focused field on the sheet in front takes the keys"
    assert facts["handoff"] == "x/b", "moving focus must commit the field left"
    assert facts["keysFreed"] == "True"
    assert facts["strayFreed"] == "True", "a key with nothing focused must free the keys"
    assert facts["behindFree"] == "True", "a field on a sheet behind must not take the keys"
    assert facts["frontTakes"] == "True", "bringing its sheet forward must take them"
    assert facts["secondOpen"] == "True/True", (
        "a Confirm opened from a dialog's OK must stay open"
    )
    assert facts["focusBack"] == "b", "focus must return past both dialogs"
    assert facts["hiddenCommits"] == "hi/True", "a focused field that hides must commit"
    assert facts["depStays"] == "True", "removing an anchor must leave dependents in place"
    assert facts["atReplaces"] == "True", "At must replace a Below placement"
    assert facts["idEnds"] == (
        "componentId may not start or end with an underscore."
    )
    assert facts["hangulWord"] == "True", "Hangul must count as word characters"
    assert facts["heldKeepsApp"] == "True", "a held close must keep the apps mounted"
    assert facts["clickResumes"] == "1", "a click after a held close must run"


def test_follow_catches_up(run_core):
    facts = parse_transcript(run_core("TestFollowCatchesUp"))
    assert facts["typingKept"] == "Ann", "a focused field keeps its text"
    assert facts["takenOnLeave"] == "Bob", "the field takes the code write once left"
    assert facts["notYet"] == "0"
    assert facts["takenOnLoad"] == "2", "a select takes the value once its item loads"
    assert facts["listenerRewrite"] == "3/3", (
        "a listener's rewrite of the key must move the control that wrote it"
    )
    assert facts["keyCase"] == "100", "WritesTo and BindValue keys match in any case"
    assert facts["sharedAfterRender"] == "2", (
        "a pick among items sharing a value must survive Render"
    )
    assert facts["snapBackQuiet"] == "5/0", "a digit that snaps back must fire nothing"


def test_bindings_and_windows(run_core):
    facts = parse_transcript(run_core("TestBindingsAndWindows"))
    assert facts["nullText"] == "Name: ", "BindText must show Null as nothing"
    assert facts["nullHidden"] == "True", "BindVisible must read Null as False"
    assert facts["textTrue"] == "True", "text that holds something reads as True"
    assert facts["nullFalse"] == "True", "BindEnabled must read Null as False"
    assert facts["navigateCase"] == "True", "Navigate must take an app id in any case"


def test_theme_builders(run_core):
    facts = parse_transcript(run_core("TestThemeBuilders"))
    assert facts["tokens"] == "True", "each theme builder must set the tokens it names"
    assert facts["stockUntouched"] == "True", "building on ThemeLight must not change it"


def test_error_words(run_core):
    facts = parse_transcript(run_core("TestErrorWords"))
    assert facts["componentNamed"] == (
        "Component 'zone': RestrictToItems applies to ComboBox only."
    )
    assert facts["kindClash"] == (
        "Component 'dark' already exists as kind Toggle, not Button."
    )
    assert facts["roleNamed"] == "Text applies to a component; this value is an app."
    assert facts["relNamed"] == (
        "Component 'tag': RightOf names 'nowhere', which is not a component of this app."
    )
    assert facts["fieldOnly"] == (
        "Component 'dark': Clearable applies to a TextInput or ComboBox."
    )


def test_hotkey_lifecycle_and_version(run_core):
    facts = parse_transcript(run_core("TestHotKeyLifecycle"))
    assert facts["procCallable"] == "True"
    assert facts["unmountClean"] == "True"
    assert facts["version"] == "1.0.2"
