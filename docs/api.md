# ReDim API

Everything lives on `ReDimUI`, a predeclared role-tagged class in the ROneCOne style. The same
type is the factory, an app, a component, a theme, an async op, and a job; each value answers
only the members of its role and raises a clear error otherwise.

## Factory

| Member | Purpose |
|---|---|
| `ReDimUI.Mount(sheet, appId)` | Create or fetch the app bound to a worksheet. Idempotent. |
| `ReDimUI.App(appId)` | Fetch a mounted app; raises if missing. |
| `ReDimUI.HasApp(appId)` | Existence probe. |
| `ReDimUI.Sender`, `SenderApp`, `SenderId` | Click context, valid inside handlers. |
| `ReDimUI.PumpOnce` | One deterministic pump tick (tests, debugging). |
| `ReDimUI.PinPumpCursor pinOn` | Opt-in steady arrow cursor while the pump is armed. Off by default so interactive shapes keep their hover hand; turn on if busy-cursor flicker is visible on your hardware. |
| `ReDimUI.AutoPump pumpOn` | Turn the wall-clock timer off for deterministic runs. |
| `ReDimUI.Shutdown` | Kill the pump and forget every app, with its window registration and the back stack. Shapes stay. |
| `ReDimUI.ThemeLight`, `ThemeDark`, `ThemeHighContrast` | Theme presets; customize with `WithPrimary`, `WithFont`. High contrast is white and yellow on black, and every pairing the controls draw passes WCAG AA. |
| `theme.ContrastReport` / `ReDimUI.ContrastRatio(foreRgb, backRgb)` | The report lists every color pairing the controls draw with its WCAG ratio, what it needs (4.5:1 for text, 3:1 for edges and the accent), and pass or fail; `ContrastRatio` computes one pair. |
| `ReDimUI.ReduceMotion motionOff` / `ReDimUI.MotionReduced` | Reduced motion follows the Windows "Show animations" setting; `True` or `False` overrides it and no argument follows Windows again. Reduced, toasts appear, move, and leave without sliding or fading. |

App ids use letters and digits only. Component ids may add single underscores.

## Windows and navigation

Sheets as forms: an app registered with `AsWindow` becomes a window, and navigation shows one
window at a time among the registered set.

| Member | Purpose |
|---|---|
| `ui.AsWindow` | Registers the app's sheet as a window. |
| `ReDimUI.Navigate appId` | Shows the target window's sheet, activates it, then very-hides every other registered window. Sheets that are not windows are never touched. |
| `ReDimUI.NavigateBack` | Pops the back stack; returns False when empty. |
| `ReDimUI.ActiveWindowId` | The currently shown window's app id. |
| `ui.OnShow "Module.Proc"` / `ui.OnHide "Module.Proc"` | Lifecycle hooks fired after navigation shows or hides the window. |
| `component.NavigatesTo "appId"` | One-declaration nav link: navigates after any OnClick handler. |
| `ui.WindowTitle "Home"` | Display name used by navigation chrome; the app id is the fallback. |
| `ui.NavBar left, top, tabWidth, tabHeight` | One tab per registered window across the top of this sheet, active tab highlighted. Bars refresh on every Navigate, picking up late registrations and pruning removed windows. |

The target is shown before others hide, because Excel requires one visible sheet at all times.
Very-hidden windows cannot be unhidden from the tab bar, and their pumps keep running: a
background window's feeds continue loading while another window is on screen. The Navigator
demo is the working reference.

## App

Component factories, get-or-create by id: `Button`, `Label`, `Card`, `Spinner`, `ProgressBar`,
`Toggle`, `TickBox`, `RadioGroup`, `Stepper`, `SlideBar`, `SelectBox`, `ComboBox`,
`TextInput`. Every control is drawn from shapes and fully themed; there are no native form
controls in the framework. Also:

| Member | Purpose |
|---|---|
| `SetState key, value` / `State(key)` / `StateOrDefault(key, fallback)` / `HasState(key)` | The store. |
| `SetStateDefault key, value` | Sets only when the key has no value; the right form for initial values. |
| (persistence) | The state store is deliberately in-memory and session-scoped; ReDim ships no persistence. Durability belongs to the host application: read the store with `State`/`StateOrDefault`, save wherever fits (a hidden sheet, workbook names, a file), and reseed on build with `SetStateDefault`, which never clobbers a value already in play. `ROneCOne.Json.Serialize`/`Deserialize` are available if JSON is the format of choice. |
| `HotKey keyCode, "Module.Proc"` / `ClearHotKeys` | Application.OnKey with cleanup on Unmount and Shutdown. |
| `OnStateChanged key, "Module.Proc"` | Zero-argument listener runs after the key changes. |
| `BeginUpdate` / `EndUpdate` | Batch several changes into one flush. |
| `Render` | Mark everything dirty and paint. Call once after building the UI. |
| `SetTheme theme` | Restyle every component, and repaint the canvas background if `PrepareCanvas` painted one. |
| `PrepareCanvas` | Paint the sheet background and hide gridlines. |
| `Toast messageText, ttlMs` | Transient card on a rail beside the content, clamped into the visible viewport, with a close button. Without `ttlMs` it stays long enough to read: three seconds plus 60 ms a character, from four seconds to twelve. `Primary`, `Success`, `Warning`, or `Danger` on the returned toast gives it an info, success, warning, or error tone, an icon and a matching edge; `.Action "Undo", "Module.Proc"` adds a button that dismisses the toast and runs the handler, and keeps the toast four seconds longer. Toasts slide up on entrance, and when one leaves the survivors slide up to fill its slot; modal chrome never shifts the rail. |
| `ToastTray rangeAddress` | Pins the tray's top-left to a range, exactly and unclamped. |
| `Confirm titleText, messageText, okProc, cancelProc, okText, cancelText` | Shapes-based modal. It takes keyboard focus: Enter confirms, Esc cancels, and Tab stays among its buttons. |
| `FocusFirst` / `DefaultButton componentId` | Keyboard focus to the first control in Tab order; the button Enter clicks from controls that do not use Enter themselves. |
| `CloseModal` | Hide the modal set. |
| `Async(opId)` / `CancelAsync opId` / `AsyncError(opId)` | Async ops (see [async.md](async.md)). |
| `Job(jobId)` / `CancelJob jobId` | Chunked or paced background work. |
| `OnError "Module.Proc"` | One-argument sink for swallowed handler failures. |
| `TickFaultCount` / `ResetTickFaults` (on `ReDimUI`) | The pump and key dispatch trap errors silently by design; every trapped fault increments this factory counter. Read it in tests or diagnostics to prove a run was clean, reset it to scope a measurement. |
| `ProtectSurface protectOn, allowCellSelection` | Opt-in app-sheet protection (UserInterfaceOnly): users cannot enter cell edit mode or drag shapes there, framework writes keep working, cell-anchored TextInput cells stay editable, and float fields type normally since they never enter cell edit. By default locked canvas cells are also unselectable - no selection rectangle on the app surface, no protected-cell warnings for stray keys - while unlocked TextInput cells stay selectable; pass `allowCellSelection:=True` to keep the whole grid selectable. OnKey capture (fields, HotKey) is unaffected. Protection state does not persist across reopen, so builds should call `ProtectSurface False` first and `ProtectSurface` after Render, as every demo does. Unmount unprotects. |
| `Unmount deleteShapes` | Remove components (and shapes) and forget the app. |

## Component builders

All fluent, all return the component:

- Geometry: `At("B2:D3")` anchors to a range and follows column widths on re-render;
  `AtRect(left, top, width, height)` uses points; `Below(otherId, gap)` and
  `RightOf(otherId, gap)` place relative to another component with `Sized(w, h)` for
  dimensions. Circular relative chains raise a clear error.
- Content: `Text`, `FontSize`, `Bold`, `BusyText`.
- Style: `Primary`, `Secondary`, `Success`, `Warning`, `Danger`, `Fill(color)`,
  `TextColor(color)`. `Warning` is amber, dark on light surfaces and bright on dark ones
  (`theme.WarningColor`).
- Visibility: `Visible(flag)`, `Enabled(flag)`.
- Accessibility: `AltText(text)` replaces the alternative text ReDim writes on the
  control's shape (see [Accessibility](#accessibility)); an empty string restores it.
- Values: `Value(number)` (progress, slider, picker index), `Checked(flag)`,
  `SliderRange(min, max, step)`.
- Item lists (`SelectBox`, `ComboBox`, `RadioGroup`, `TransferList`, `CheckList`):
  `Items("A", "B", ...)` replaces; `ItemsFrom(source)` replaces from a 1D array, a
  Collection, a Range (one item per non-empty cell), or a ROneCOne sequence;
  `AddItem(text, atPosition)` appends or inserts; `RemoveItem(indexOrText)`; `ClearItems`;
  read back with `ItemCount` and `ItemTextAt(position)`. The selected item survives inserts
  and unrelated removals; removing it clears the selection to the placeholder.
  Programmatic mutations re-render but do not write `WritesTo` state or fire `OnChange`;
  those belong to user interaction and explicit `SetState`.
- Placeholder (`SelectBox`): `Text` is the placeholder the face shows while nothing is
  selected, and `Value(0)` clears the selection back to it. Like the item mutations, the
  clear writes no `WritesTo` state and fires no `OnChange`.
- Drop lists (`SelectBox`, `ComboBox`): `ListRows(n)` sets how many item rows the open
  list shows at once, eight by default. A longer list windows behind clickable pager rows
  at its edges, each showing an arrow and the count of items beyond it. The current item's
  row carries a check in a gutter every row shares, so the selection reads without color.
  An open list with nothing to show says so in an inert row, "No matches" under a combo's
  filter or "No items" in an empty select. A list that would run past the bottom of the
  visible window opens upward when there is more room above. One list is open per app:
  opening one closes the others, and so does a click on another control, a press anywhere
  off the face and rows, or a move of the grid selection.
- Bindings: `BindText(key, template)` where `{0}` is the value, `BindValue(key)`,
  `BindVisible(key, invert)`, `BindEnabled(key, invert)`, `WritesTo(key)`. The invert flag
  serves the disable-while-busy pattern: `BindEnabled "anyRunning", True`.
- Behavior: `OnClick "Module.Proc"`, `OnClickAsync "Module.Proc"`, `OnChange "Module.Proc"`.
- Keyboard: `TabIndex(n)` orders Tab, `Focus` gives the control keyboard focus,
  `AccessKey(letter)` binds Alt+letter for a button or tick box, and `Clearable` gives a
  float TextInput or ComboBox a clear button (see
  [Keyboard focus for every control](#keyboard-focus-for-every-control)).
- Text fields (float `TextInput` and `ComboBox`): `Placeholder(hint)`, `OnInput(proc)` with
  `DebounceMs(ms)`, `Numeric`, `MaxLength(n)`, and `Validates(checkProc)` read back with
  `ValidationError` (see [Field rules](#field-rules)); `MultiLine` and `AutoGrow(maxLines)`
  for a TextInput and `RestrictToItems` for a ComboBox.
- Reads: `CurrentValue`, `CurrentText`, `IsChecked`, `IsEnabled`, `IsVisible`, `IsBusy`, `InputValue`
  (TextInput and ComboBox; reads the float buffer or the backing cell, and assigning it
  writes without firing change events).
- `Remove` deletes the component and its shapes.

Handlers are zero-argument public procedures referenced as `"Module.Proc"`. Inside a handler,
`ReDimUI.Sender` carries who fired and `ReDimUI.SenderPart` names the clicked sub-shape when a
composite widget fired (toggle knob, checkbox caption, select option).

## The drawn control family

Excel fixes the fonts and look of form controls, so as of 0.5.0 ReDim draws every control from
shapes and ships no native form controls at all. Full styling control, pure shapes, no added
dependencies:

- `TickBox`: themed box, check glyph, caption; box and caption both toggle on click.
- `RadioGroup`: single-select option rows, a control native form controls never offered.
- `Stepper`: numeric entry as minus and plus around a value face, honoring `SliderRange` -
  the precise keyboard-free form of numeric input.
- `SlideBar`: a drawn slider with true press-drag, without blocking. Shape OnAction only
  fires at mouse up, so the pump's frames watch for the left-button press edge themselves,
  hit-test the cursor against the track, and run the drag session: the value follows the
  cursor with snapped live state writes while the button is held (thumb shown in the accent
  color), `OnChange` fires once at release if the value moved, and the release click Excel
  then delivers is swallowed. A plain tap sets the value at the press point. Loop-free by
  construction: one key-state poll and at most one cursor read per 16 ms frame, and every
  other pump duty keeps running mid-drag. A visible slider on the active sheet keeps the
  pump armed so presses are never missed - the idle cost is that one poll per frame.
  Coordinates come from `GetCursorPos` plus a DPI-and-zoom-aware inversion of
  `PointsToScreenPixels`; frozen-pane splits skew that calibration, so keep app surfaces
  unsplit. Floating chrome claims the points it covers, so a press on an open drop list
  or a modal overlay never reaches a track painted underneath. Use `Stepper` for
  precision.
- `SelectBox`: a themed face, caret, and option list in place of the native dropdown.
  `Text` is the placeholder shown while nothing is selected. The list windows to
  `ListRows` rows (eight by default): opening scrolls the selection into view, and
  clickable pager rows at the list edges (arrow plus the count beyond that edge) page the
  window, appearing only when something lies beyond them. Picking a new item writes the
  `WritesTo` state and fires `OnChange`; re-picking the selected item only closes the
  list, the same rule `RadioGroup` follows for its selected row.
- `TransferList`: a dual listbox - two panels with counted headers, selectable rows, and
  four move buttons (`>`, `>>`, `<`, `<<`). `Items`/`ItemsFrom` and the item APIs feed the
  available side, `ChosenFrom` seeds the chosen side, `Captions` names the headers, and
  `ChosenCount`/`ChosenTextAt` read the result. Rows multi-select by click-to-toggle:
  each plain click adds or removes that row from the panel's selection, and `>`/`<` move
  every selected row in list order. (Ctrl+click cannot exist on a drawn control - Excel
  reserves modifier-clicks on macro shapes for selecting the shape itself, the same rule
  behind the design-time escape hatch below.) `WritesTo` carries the chosen items joined
  with a comma and space; `OnChange` fires once per user transfer, and selecting rows
  fires nothing. A selected row shows a check as well as the accent fill. Rows render up
  to the panel's height; when a list outgrows its panel, paging arrows appear on the
  panel's right edge, 18-point targets, and move the window a page at a time, and the
  header counts stay honest about totals.
- `ComboBox`: an editable combo with a caret and a filtered drop list, sharing the item
  APIs. Place it with `AtRect` (or `Below`/`RightOf`) and it is a float field: click to
  focus, type, and the list re-filters on every keystroke. Anchoring to a cell with `At`
  keeps the 0.8.0 cell-backed mode, where Excel's edit-mode VBA pause limits filtering to
  commit moments; its face covers the cell, so a click on it selects the cell for typing
  as well as opening the list. In both modes a pick, by click or by Enter on the
  highlighted row, writes the value and closes the list; when it changes the value it
  also writes the `WritesTo` state and fires `OnChange`, exactly once. Re-picking the
  value already there only closes the list. For a float combo, the value a pick replaces
  is the text the field held when focus arrived, so text typed only to filter the list
  never counts. The drop list windows to `ListRows` rows (eight by default): Up and Down
  walk a highlight that scrolls the window, Enter takes the highlighted match, typing
  re-filters, and clickable pager rows at the list edges (arrow plus the count beyond that
  edge) page the window for the mouse, appearing only when something lies beyond them.
  When the text names an item, as it does after a pick, the list reopens unfiltered and
  scrolled to that item, and Down walks on from it; the first edit filters again. While
  the text filters the list, each row bolds what the text matched, and the first item
  that begins with the text shows the rest of its name after the caret in muted ink.
  Right at the end of the text, or Tab, takes it in the item's spelling; Enter commits
  what was typed. A click on a focused combo's text places the caret, and the arrow at
  its right edge toggles the list. `RestrictToItems` limits a float combo's commits to
  its items (see [Field rules](#field-rules)).
- `TextInput`: a text field. Float by default (`AtRect`), cell-backed with `At` when you
  want the value to live in the grid. A cell-backed field's frame covers its cell, so a
  click on the frame selects the cell: typing replaces the value and F2 edits it in place,
  and the click itself fires no `OnChange`.
- `CheckList`: a checkbox list - one box-and-caption row per item, any number checked,
  with a select-all header on by default (`WithSelectAll False` opts out). The header box
  is tri-state (empty, check, dash for mixed), clicking it checks everything unless all
  are already checked, and its caption counts (`Select all (2/4)`). Box and caption both
  toggle their row. `CheckedFrom` seeds by text; `SetItemChecked`/`IsItemChecked`/
  `CheckedCount` are the silent programmatic surface; checks follow their items through
  `AddItem`/`RemoveItem`. `WritesTo` carries checked items joined with ", "; `OnChange`
  fires once per toggle, select-all included.
- `Image`: a picture as a control - a rounded rectangle whose fill is the picture, so it
  clicks, adopts, and snaps back like everything else, and the picture embeds in the
  workbook. `Source(path)` takes a file path (no URLs) and loads once per distinct path;
  `BindSource(key)` drives it from state. The image stretches to the declared rectangle;
  a missing source renders a themed placeholder, but an already-embedded picture is kept
  even after its source file goes away.
- `Toggle`: the pill switch for booleans. Switched on, the knob takes the theme's
  `OnPrimary` ink, as Windows draws it, so it stays visible on every accent track.

## Float fields and keyboard focus

As of 0.9.0 no control needs a cell. `TextInput` and `ComboBox` placed with `AtRect`,
`Below`, or `RightOf` render and edit entirely on their shapes: clicking a field gives it
keyboard focus with the caret where the click landed, an accent ring and an insertion bar
appear (blinking at the Windows caret
rate, or steady when Windows is set not to blink), and characters go to the
component's text buffer instead of a cell. That sidesteps Excel's edit-mode VBA pause - the
framework sees every keystroke, which is what makes live combo filtering possible - and it
means typing works on `ProtectSurface` sheets, where cell edit is locked out.

Focus mechanics, all automatic:

- One control holds keyboard focus at a time; focus moving on commits a field it leaves.
- Enter commits: the buffer becomes the value, `WritesTo` state is written and `OnChange`
  fires if the text changed, and then Enter clicks the app's `DefaultButton` if it has
  one. Tab and Shift+Tab commit the same way and move focus to the next or previous
  control in Tab order (see [Keyboard focus for every control](#keyboard-focus-for-every-control));
  with no other control to move to, they commit and leave. Tab first takes a combo's
  shown suggestion. A combo commit whose text names an item, in any case, takes that
  item. A `MultiLine` TextInput follows textarea conventions
  instead: Enter inserts a newline and keeps focus, while Tab, Ctrl+Enter, and clicking
  away commit; the value carries its newlines into state. Size the rectangle for the
  lines you expect - roughly 15 points per line plus 6 points of margin; a line that
  cannot fully fit is not drawn at all, which reads as a missing line. Or let
  `AutoGrow(maxLines)` size it: the field grows a line at a time from the height it was
  given, up to `maxLines` lines (six by default), shrinks back as lines go, and moves
  anything placed `Below` it. At rest a line too long for the field wraps, and the
  height counts the wrapped lines, so the field keeps one height with focus or without.
  `AutoGrow` turns `MultiLine` on.
- Overflow follows the caret. Shape text cannot scroll, so a focused field renders the
  tail window of its buffer - the last lines that fit (multi-line) or the rightmost
  characters that fit (single-line and combo) - with a leading ellipsis marking trimmed
  content. The buffer keeps the complete text and commits intact; unfocused fields show
  their beginning.
- Esc reverts the field to the text it had when focus arrived, fires nothing, and leaves.
  A combo with its list open closes the list on the first Esc, keeping focus and text,
  and reverts on the second. `Clearable` gives either field a clear button for emptying
  it.
- Clicking anywhere off the field commits, through two frame-driven signals. A cell click
  moves the selection, and the focused field's frames poll the selection - the press
  itself can be invisible (the grid's selection mouse loop holds timer messages until the
  button is back up), but the selection it leaves behind is durable state the next frame
  sees. Sheet navigation commits the same way, and none of it depends on application
  events. (Arrow keys no longer move the selection while a field is focused - they
  edit.) Everything that changes no selection - either mouse button on shapes, chrome,
  or other windows - is caught by the same press-edge watch that drives slider drags,
  and a click on another ReDim control commits before that control's handler runs. The one blind spot: re-clicking the already-selected cell during a starved
  moment changes nothing observable; the next keystroke, click, or frame resolves it. On a
  default `ProtectSurface` sheet the canvas is unselectable, so grid clicks move no
  selection at all - there, commit by Enter, Tab, or clicking any control, which is the
  natural flow on an app surface anyway.
- `InputValue` reads and writes the buffer in float mode, the cell in cell mode.

Capture uses `Application.OnKey`, bound only while a control holds focus and released when
it leaves, so sheet typing is untouched the rest of the time. The bound set is the
practical editing and navigation set: the letters a to z (with Shift capitals), digits,
space, punctuation and symbols (bound by character, so each follows the active keyboard
layout), Backspace, Del, the arrow keys, Home, End, Page Up, Page Down, Tab and
Shift+Tab, Enter (and Ctrl+Enter), Esc, Alt+Down, Alt+Up, F4, and the selection,
clipboard, and undo chords under [Text editing](#text-editing). Editing is full caret
editing: arrows move the insertion point, characters insert at it, Backspace and Del
delete around it, Home and End jump the line edges, and Up and Down move across hard
lines with the column clamped (a long wrapped line counts as one line). While a field is
focused the arrows belong to editing, so they do not move the cell selection. Keys outside
the bound set, accented letters and IME input among them, fall through to the grid as
usual; on a `ProtectSurface` sheet Excel answers those with its
protected-cell notice, and protection stays on the whole time - OnKey capture works fine
under protection (verified with message-level keystrokes). Apps that bind arrow HotKeys
should re-arm them after field focus sessions if they mix the two. `ReDimUI.HasKeyboardFocus`
and `ReDimUI.FocusedComponentId` report the current holder; `RdxReleaseKeys` is the panic
release that unbinds everything regardless of state.

## Text editing

A focused float field edits the way a Windows text box does.

| Keys | Action |
|---|---|
| Shift with Left, Right, Up, Down, Home, or End | Extend the selection. |
| Ctrl+Left, Ctrl+Right | Move a word; with Shift, select by words. |
| Ctrl+Home, Ctrl+End | Jump to the start or end of the text; with Shift, select to there. |
| Ctrl+Backspace, Ctrl+Del | Delete the word before or after the caret. |
| Ctrl+A | Select everything. |
| Ctrl+C, Ctrl+X, Ctrl+V | Copy, cut, and paste through the Windows clipboard, as Unicode text. |
| Ctrl+Z; Ctrl+Y or Ctrl+Shift+Z | Undo; redo. |

- The selection shows in the accent color with accent ink. Typing, a delete, or a paste
  replaces it, and Left or Right without Shift collapses it to its start or end.
- A word is a run of letters, digits, underscores, and characters past ASCII.
- A paste keeps what the field takes: a single-line field turns line breaks and tabs into
  spaces, a trailing line break (a copied cell brings one) is dropped, and `Numeric` and
  `MaxLength` apply as they do to typing.
- Undo steps back through the edits since the field took focus, up to 100 of them. Typing
  groups into one step until it pauses for a second; a delete, a cut, a paste, and a taken
  combo suggestion are steps of their own.
- A click puts the caret at the character boundary nearest the pointer, and a double click
  within the Windows double-click time selects the word under it. A click raised with the
  pointer off the face, from code for instance, places no caret: a field it focuses starts
  with the caret at the end.
- `CurrentText` and `InputValue` return the text alone, without the insertion bar.

## Field rules

Builders for float `TextInput` and `ComboBox` fields:

- `Placeholder "Search..."` shows a hint in muted ink while the field is empty, focused or
  not. The hint is not text: `InputValue` stays empty, and the alternative text includes
  the hint.
- `OnInput "Module.Proc"` runs after every edit that changes the text, typing, deletes,
  paste, cut, and undo alike. `ReDimUI.Sender` is the field, so
  `ReDimUI.Sender.InputValue` is the text so far. `DebounceMs 300` waits until typing
  pauses that long; a pending call runs before focus leaves the field. `OnChange` still
  fires once per commit.
- `Numeric` keeps digits, one decimal separator, and a leading minus. The separator is the
  locale's, and both the period and the comma type it. `Numeric allowDecimal:=False` refuses
  the separator and `Numeric allowNegative:=False` the minus.
- `MaxLength 5` caps the text: typing and pasting stop at the cap, and a count under the
  field's right edge shows the length against it.
- `Validates "Module.CheckEmail"` names a Public Function that takes the text and returns
  "" when it is valid or a message when it is not. It runs on commit. An invalid field
  gets a danger border and the message under it, and from then on every edit rechecks
  until the text passes. The commit still happens; `ValidationError` reads the message,
  so gate whatever the value feeds on it.
- `RestrictToItems` makes a float `ComboBox` commit only its items. A commit takes the
  item the text names in any case, or else the first item the text begins, in the item's
  spelling. An empty text stays empty, and any other text goes back to what the field held
  when focus arrived. A cell-backed combo's cell is Excel's to restrict, with data
  validation.

```vba
ui.TextInput("email").AtRect(24, 24, 220, 22).Placeholder "name@example.com"
ui.TextInput("email").Validates "Checks.Email"
ui.TextInput("qty").AtRect(24, 70, 80, 22).Numeric allowDecimal:=False
ui.TextInput("find").AtRect(24, 116, 220, 22).OnInput "Search.Refilter"
ui.TextInput("find").DebounceMs 250
ui.ComboBox("fruit").AtRect(24, 162, 220, 22).Items("Apple", "Banana").RestrictToItems
```

## Keyboard focus for every control

Every interactive control takes keyboard focus: buttons, toggles, tick boxes, radio groups,
steppers, sliders, selects, check lists, transfer lists, images with a click handler, and
float fields. Focus comes from the keyboard or from code: Tab and Shift+Tab walk the app's
controls, `component.Focus` and `ui.FocusFirst` place it, and a modal takes it. A click
focuses only a text field, since someone who clicks a button in Excel expects the grid to
keep the keys.

- Tab order: controls with a positive `TabIndex` come first, in ascending order, then the
  rest in creation order. `TabIndex(-1)` leaves a control out of Tab while `Focus` still
  reaches it. Tab wraps within the app. Focus that lands outside the visible window
  scrolls the control into view by rows and columns, leaving the selection where it is.
- A focused control wears an accent ring three points outside its bounds; a text field
  shows focus on its own border. A check list or transfer list also shows a dotted cursor
  on the row its arrow keys have reached.
- Focus ends on Esc, on a click on another control, on a press anywhere off the control,
  or when the grid selection moves: the watch that ends a text field's focus, for every
  kind. A control that turns hidden or disabled gives focus up.
- Enter on a control with no use for it clicks the app's default button,
  `ui.DefaultButton "save"`. A focused button takes Enter itself.
- Modals trap focus. `Confirm` puts focus on OK, Tab stays among the dialog's buttons,
  Enter confirms, Esc cancels (or confirms when there is no Cancel), and closing the
  dialog returns focus to the control that had it.
- Access keys: `AccessKey "s"` on a button or tick box underlines the first matching
  character of its text and clicks the control on Alt+S while its sheet is in front. The
  chords are bound with `Application.OnKey` only while such a sheet is active, and while
  bound they take that Alt+letter from Excel, so pick letters your users do not need for
  the ribbon. An ampersand in the text stays literal.

| Control | Keys |
|---|---|
| Button, Image | Space or Enter clicks. |
| Toggle, TickBox | Space toggles. |
| RadioGroup | Arrows move the selection, wrapping at the ends; Home and End jump. Each move fires `OnChange`. |
| Stepper | Up and Right step up, Down and Left step down, Page Up and Page Down step ten times, Home and End jump to the range ends. |
| SlideBar | Arrows move a step, Page Up and Page Down a tenth of the range in whole steps, Home and End go to the ends. |
| SelectBox | Closed: arrows, Home, End, Page Up, and Page Down change the selection, and Space, Alt+Down, or F4 opens the list. Open: they move the highlight; Enter, Space, or Alt+Up takes it, Tab takes it and moves on, and Esc or F4 closes. Letters jump to the next item that starts with them, open or closed: letters typed within a second build a prefix, and one letter typed again steps through its items. |
| CheckList | Up and Down move the row cursor, the select-all header included; Home and End jump; Space toggles the cursor's row. |
| TransferList | Up and Down move the row cursor, Left and Right switch panels, Space toggles the cursor's row in the selection, and Enter moves the panel's selection across, or the cursor's row when nothing is selected. |
| TextInput, ComboBox | The editing keys under [Text editing](#text-editing); a combo also opens with Alt+Down or F4, closes with Alt+Up, pages its list with Page Up and Page Down, and takes its suggestion with Right at the end of the text or Tab. |

## Accessibility

- Alternative text: ReDim writes a description on each control's shape from its kind, text,
  and state, such as "Save, button", "Agree, checkbox, checked", "Drop-down, South",
  "Progress, 40 percent" (in 5 percent steps), with ", unavailable" when disabled. Labels,
  cards, and toasts carry none, since their text is what a reader announces. `AltText`
  replaces the description.
- Keyboard: every control works without a mouse, and focus shows as a ring or a field
  border (see [Keyboard focus for every control](#keyboard-focus-for-every-control)).
- State without color: a drop list's current item and a transfer panel's selected rows
  show a check, not only a fill.
- Contrast: `ThemeHighContrast` passes WCAG AA for every pairing the controls draw, and
  `theme.ContrastReport` checks any theme, a custom one included. The light and dark
  presets keep their look, and their reports show one shortfall each: the field edge
  (`Border` on `Surface`) reaches 1.32:1 and 2.01:1 against the 3:1 non-text minimum.
- Motion: toasts stop sliding and fading when Windows animations are off or
  `ReduceMotion True` is set. The spinner keeps turning; it is status, not decoration.
- Targets: transfer paging arrows and the toast close button are 18 points square, the
  WCAG 2.5.8 minimum of 24 pixels at 96 DPI.

## Shapes are framework-owned

Every ReDim shape carries the dispatcher as its `OnAction`, including kinds with no click
behavior. A shape with a macro assigned runs it instead of being selected, so plain clicks
cannot drag a modal card, label, or progress bar out of position. For deliberate design-time
manipulation, Ctrl+click selects a shape as usual; geometry diffs against the live shape, so
the next `Render` (or the next flush touching that component) snaps it back to its declared
rectangle.

## Statement-form chaining rule

VBA accepts a fluent chain as a statement only when the final call uses bare arguments or a
single parenthesized argument:

```vba
ui.Spinner("busy").AtRect 400, 18, 26, 26        ' right
ui.Spinner("busy").AtRect(400, 18, 26, 26)       ' compile error: Syntax error
ui.Label("x").At("B2").Text("hello")             ' fine: one argument
```

Mid-chain calls are expression position and may keep their parentheses. The repository's compile
gate (`tests/python/test_compile.py`) catches violations that static analysis cannot.

## Case-insensitive shadowing rule

VBA identifiers are case-insensitive. A local `Dim ownerApp` shadows a property named
`OwnerApp`, so `Set ownerApp = OwnerApp` self-assigns Nothing. Never name a local after a member.

## Identifier casing

The same case-insensitivity leaves each project with one spelling per name, and a
declaration in any module can set it for all of them. A parameter named `value` anywhere
in the project turns every `.Value` into `.value`, and the next module export carries
that into version control as noise. ReDim declares no name in a casing that differs from
the Excel, Office, VBA, or stdole type libraries or from its own members, and
`tests/python/test_casing.py` holds the runtime and the demos to that, ending with a VBE
export round trip. Host code stays quiet the same way: the samples name their app
variable `ui`, because a variable named `app` collides with `ReDimUI.App`.
