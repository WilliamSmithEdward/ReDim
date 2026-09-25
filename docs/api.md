# ReDim API

Everything lives on `ReDimUI`, a predeclared role-tagged class in the ROneCOne style. The same
type is the factory, an app, a component, a theme, an async op, and a job; each value answers
only the members of its role and raises a clear error otherwise, such as "Text applies to a
component; this value is an app." A component's errors start with its id, as in
"Component 'zone': RestrictToItems applies to ComboBox only.", so a failure in a long build
names the control that raised it.

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
| `ReDimUI.ThemeSystem` | The Windows look: `ThemeDark` while Windows apps use dark mode and `ThemeLight` otherwise, with the Windows accent color as the primary. The accent moves toward black on the light theme, or white on the dark one, until it clears 3:1 against the surface and 4.5:1 under its ink. Pair it with `ui.FollowSystemTheme` to keep up with changes. |
| `ReDimUI.IconGlyph(name)` / `IconNames` / `IconFont` | The character that draws a named icon, for text of your own; the list of names (see [Icons](#icons)); and the Windows icon font they draw in, Segoe Fluent Icons where Windows 11 installed it and Segoe MDL2 Assets otherwise. |
| `theme.ContrastReport` / `ReDimUI.ContrastRatio(foreRgb, backRgb)` | The report lists every color pairing the controls draw with its WCAG ratio, what it needs (4.5:1 for text, 3:1 for edges and the accent), and pass or fail; `ContrastRatio` computes one pair. |
| `ReDimUI.ReduceMotion motionOff` / `ReDimUI.MotionReduced` | Reduced motion follows the Windows "Show animations" setting; `True` or `False` overrides it and no argument follows Windows again. Reduced, toasts appear, move, and leave without sliding or fading, and knobs and tab bars land without gliding. |
| `ReDimUI.Version` | The runtime's version, such as `"1.0.0"`. |
| `ReDimUI.SetUIText key, words` / `UIText(key)` / `UITextKeys` / `ResetUIText` | The words ReDim draws and announces on its own, for an app in another language (see [UI text](#ui-text)). |

A theme answers its tokens, for shapes of your own that should match the controls:
`PrimaryColor`, `OnPrimaryColor`, `SurfaceColor`, `OnSurfaceColor`, `MutedColor`,
`OnMutedColor`, `SuccessColor`, `DangerColor`, `WarningColor`, `BorderColor`, `CanvasColor`,
`FontName`, and `BaseFontSize`, as in `ui.Theme.PrimaryColor`.

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
`Skeleton`, `Toggle`, `TickBox`, `RadioGroup`, `Stepper`, `SlideBar`, `SelectBox`, `ComboBox`,
`TransferList`, `CheckList`, `TextInput`, `DatePicker`, `Image`, `Tabs`, `Table`, `Badge`,
`Expander`, `MenuButton`, `Sparkline`, and the `Stack` layout container (see
[Layout](#layout)). Every
control is drawn from shapes and fully themed; there are no native form controls in the
framework. Also:

| Member | Purpose |
|---|---|
| `SetState key, value` / `State(key)` / `StateOrDefault(key, fallback)` / `HasState(key)` | The store. |
| `StateKeys` | Every key in the store once, in the order it was first set, as an array from 0, whether code or a control's `WritesTo` set it. |
| `SetStateDefault key, value` | Sets only when the key has no value; the right form for initial values. |
| (persistence) | The state store is deliberately in-memory and session-scoped; ReDim ships no persistence. Durability belongs to the host application: walk the store with `StateKeys` and `State`, save wherever fits (a hidden sheet, workbook names, a file), and reseed on build with `SetStateDefault`, which never clobbers a value already in play. `ROneCOne.Json.Serialize`/`Deserialize` are available if JSON is the format of choice. |
| `HotKey keyCode, "Module.Proc"` / `ClearHotKeys` | Application.OnKey with cleanup on Unmount and Shutdown. |
| `OnStateChanged key, "Module.Proc"` | Zero-argument listener runs after the key changes. |
| `BeginUpdate` / `EndUpdate` | Batch several changes into one flush. |
| `Render` | Mark everything dirty and paint. Call once after building the UI. |
| `FlushDirty` | Paint the components changed since the last paint, now. A change outside `BeginUpdate` paints on its own, so code rarely needs it. |
| `AppId` / `Sheet` / `Theme` / `Component(id)` / `ComponentCount` | Readers: the app's id, its worksheet, its theme, a component by id (raises when there is none), and how many it holds. |
| `IsWindow` / `IsSurfaceProtected` | Whether `AsWindow` registered the app, and whether `ProtectSurface` is on. |
| `SetTheme theme` | Restyle every component, and repaint the canvas background if `PrepareCanvas` painted one. |
| `PrepareCanvas` | Paint the sheet background and hide gridlines. |
| `Toast messageText, ttlMs` | Transient card on a rail beside the content, clamped into the visible viewport, with a close button. A top toast too tall for the room below the rail lifts the rail, and a toast shown while its sheet is behind another fits to the window once the sheet is in front. Without `ttlMs` it stays long enough to read: three seconds plus 60 ms a character, from four seconds to twelve. The card is 240 points wide and as tall as its wrapped message, up to 160 points, where the message ends in an ellipsis. `.MinWidth` and `.MaxWidth` on the returned toast let it fit its words between the two, wrapping past the wider (a `MinWidth` under 240 lets a short message make a narrow toast), and `.MaxHeight` moves the cap. `Primary`, `Success`, `Warning`, or `Danger` gives it an info, success, warning, or error tone, an icon and a matching edge; `.Action "Undo", "Module.Proc"` adds a button that dismisses the toast and runs the handler, and keeps the toast four seconds longer, and `.ActionBorder rgb` colors that button's border, the accent until set. `.OnClick "Module.Proc"` runs when the card itself is clicked, with the toast as `ReDimUI.Sender`. Any click puts the toast away at once, and the close button runs no handler. A toast under the pointer stops counting down and keeps at least a second once the pointer leaves. Toasts stack by their heights, slide up on entrance, and when one leaves the survivors slide up to fill its place; modal chrome never shifts the rail. |
| `ToastTray rangeAddress` | Pins the tray's top-left to a range, exactly and unclamped. |
| `Confirm titleText, messageText, okProc, cancelProc, okText, cancelText` | Shapes-based modal. It takes keyboard focus: Enter confirms, Esc cancels, and Tab stays among its buttons. The buttons read the `OK` and `Cancel` [UI text](#ui-text) unless given words; `cancelText:=""` leaves Cancel out. |
| `FocusFirst` / `DefaultButton componentId` | Keyboard focus to the first control in Tab order; the button Enter clicks from controls that do not use Enter themselves. |
| `ValidateAll` | Checks the form before a save: every enabled float field with `Required` or `Validates`, the way a commit does, whether it shows or sits on a Tabs or Expander panel not shown. Hidden and disabled fields are skipped. Each field that fails shows its message, and the first in Tab order takes focus, its tab turned to or its expander opened first. True when every field passes (see [Field rules](#field-rules)). |
| `FocusRings ringsOn` | An accent ring around a focused control that is not a text field (see [Keyboard focus for every control](#keyboard-focus-for-every-control)). Off by default. |
| `PointerEffects effectsOn` | Hover and press looks for the app's controls, and an open list whose highlight follows the pointer (see [Pointer](#pointer)). Off by default. |
| `AddCommand commandWords, handlerProc, iconName` | Adds an entry to the app's command palette: its text, the Public procedure it runs, and an optional icon. Inside the handler, `ReDimUI.SenderApp.LastCommand` names the entry. On a `MenuButton` the same builder adds a menu row. |
| `CommandPalette paletteKey` / `OpenCommandPalette` | Turns on the command palette and binds its key, Ctrl+Shift+P unless another OnKey code is given (`""` binds none); `OpenCommandPalette` opens it from code (see [Command palette](#command-palette)). |
| `FollowSystemTheme followOn` | The app takes `ReDimUI.ThemeSystem` now, and again whenever Windows switches between light and dark mode or changes its accent. The app notices as its sheets activate, the selection moves, or the pump runs, reading the registry at most every two seconds. Called before `Render`, it sets the theme without rendering early. `False` stops following and keeps the theme in place. |
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
  dimensions. Circular relative chains raise a clear error. `InStack(stackId)` hands the
  position to a `Stack` (see [Layout](#layout)).
- Content: `Text`, `FontSize`, `Bold`, `BusyText` (what a busy button shows, the
  `Working` [UI text](#ui-text) unless given; `BusyText ""` keeps the button's text).
- Style: `Primary`, `Secondary`, `Success`, `Warning`, `Danger`, `Fill(color)`,
  `TextColor(color)`. `Warning` is amber, dark on light surfaces and bright on dark ones
  (`theme.WarningColor`).
- Visibility: `Visible(flag)`, `Enabled(flag)`. `OnTab(tabsId, tabNumber)` puts the
  control on a tab's panel of a `Tabs` control, where it shows only while that tab does;
  `OnTab "", 0` takes it off.
- Accessibility: `AltText(text)` replaces the alternative text ReDim writes on the
  control's shape (see [Accessibility](#accessibility)); an empty string restores it.
- Icons (`Button`, `Label`): `Icon("Save")` draws a Windows icon before the text, or
  alone when there is none (see [Icons](#icons)).
- Badges (any control): `BadgeText("3")` puts a count or short note in a danger-colored
  pill on the control's top-right corner, such as the unread count on an Inbox button;
  it hides with the control, a click on it acts as a click on the control, and `""`
  takes it off.
- Tooltips: `Tooltip(text)` shows a note once the pointer rests on the control, and
  `DisabledReason(text)` says why a disabled control is disabled (see [Pointer](#pointer)).
- Values: `Value(number)` (progress, slider, picker index), `Checked(flag)`,
  `SliderRange(min, max, step)`, which refuses a maximum at or below the minimum and a step
  of zero or less.
- Item lists (`SelectBox`, `ComboBox`, `RadioGroup`, `TransferList`, `CheckList`):
  `Items("A", "B", ...)` replaces; `ItemsFrom(source)` replaces from a 1D array, a
  Collection, a Range (one item per non-empty cell), or a ROneCOne sequence;
  `AddItem(text, atPosition)` appends or inserts; `RemoveItem(indexOrText)`; `ClearItems`;
  read back with `ItemCount` and `ItemTextAt(position)`. The selected item survives inserts
  and unrelated removals; removing it clears the selection to the placeholder. A
  `RadioGroup` or `CheckList` left with no items shows nothing until items return.
  Programmatic mutations re-render but do not write `WritesTo` state or fire `OnChange`;
  those belong to user interaction and explicit `SetState`.
- Item values: an item can carry a value apart from its text, `AddItem("Medium", , 20)` or
  `ItemsFrom(texts, values)` with the values in a source of the same shape (a range of
  items with gaps pairs with its range of values cell by cell). A pick then writes the
  value to `WritesTo` instead of the text, with the value's own type, and a check list or
  transfer list joins the values of its checked or chosen items. `ItemValueAt(position)`
  and `ChosenValueAt(position)` read a value, or the text for an item without one. Values
  belong to item texts, without case, so they follow items between transfer panels and
  two items with one text share a value. A combo's free text writes the text. Replacing
  the list with `Items`, `ItemsFrom`, or `ClearItems` drops its values; a transfer list's
  chosen items keep theirs, and `ChosenFrom(texts, values)` sets them.
- Placeholder (`SelectBox`): `Text` is the placeholder the face shows while nothing is
  selected, and `Value(0)` clears the selection back to it. Like the item mutations, the
  clear writes no `WritesTo` state and fires no `OnChange`.
- Drop lists (`SelectBox`, `ComboBox`): `ListRows(n)` sets how many item rows the open
  list shows at once, eight by default. A longer list windows between clickable pager rows
  at its edges, each showing an arrow and the count of items beyond it. Both pagers stay
  in place at every page, so clicking one spot pages through the whole list; at the first
  or last page the pager with nothing beyond it shows its arrow alone, muted, and a click
  on it does nothing. The current item's
  row carries a check in a gutter every row shares, so the selection reads without color.
  An open list with nothing to show says so in an inert row, "No matches" under a combo's
  filter or "No items" in an empty select. A list that would run past the bottom of the
  visible window opens upward when there is more room above. One list is open per app:
  opening one closes the others, and so does a click on another control, a press anywhere
  off the face, rows, and pagers once the button comes back up, or a move of the grid
  selection. A click Excel delivers on the list itself after such a press keeps it open,
  so a pointer reading that lands just off the list never closes it under a click on
  its own pager.
- Bindings: `BindText(key, template)` where `{0}` is the value, `BindValue(key)`,
  `BindVisible(key, invert)`, `BindEnabled(key, invert)`, `WritesTo(key)`. The invert flag
  serves the disable-while-busy pattern: `BindEnabled "anyRunning", True`. A `Toggle`,
  `TickBox`, or `Expander` bound with `BindValue` follows `True` and `False`, so
  `WritesTo "darkMode"` with `BindValue "darkMode"` keeps a switch and its state in step
  whichever side changes.
- Behavior: `OnClick "Module.Proc"`, `OnClickAsync "Module.Proc"`, `OnChange "Module.Proc"`.
- Keyboard: `TabIndex(n)` orders Tab, `Focus` gives the control keyboard focus,
  `AccessKey(letter)` binds Alt+letter for a button or tick box, `Shortcut(keyCode)` binds
  a keyboard shortcut such as `"^s"` for Ctrl+S that clicks the control, and `Clearable`
  gives a float TextInput or ComboBox a clear button (see
  [Keyboard focus for every control](#keyboard-focus-for-every-control)).
- Text fields (float `TextInput` and `ComboBox`): `Placeholder(hint)`, `OnInput(proc)` with
  `DebounceMs(ms)`, `Numeric`, `MaxLength(n)`, `Required`, `ErrorText(message)`, and
  `Validates(checkProc)` read back with `ValidationError` (see [Field rules](#field-rules));
  `MultiLine`, `AutoGrow(maxLines)`, and `Masked` for a TextInput and `RestrictToItems`
  for a ComboBox.
- Adornments (any control): `Caption(text)` puts a label in small text above the control,
  and `Hint(text)` helper text in muted ink below it (see [Field rules](#field-rules)).
- Reads: `ComponentId`, `CurrentValue`, `CurrentText`, `IsChecked`, `IsEnabled`,
  `IsVisible`, `IsBusy`, `InputValue` (TextInput and ComboBox; reads the float buffer or
  the backing cell, and assigning it writes without firing change events).
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
  the precise keyboard-free form of numeric input. At the minimum the minus reads muted,
  and at the maximum the plus.
- `SlideBar`: a drawn slider with true press-drag, without blocking. Shape OnAction only
  fires at mouse up, so the pump's frames watch for the left-button press edge themselves,
  hit-test the cursor against the track, and run the drag session: the value follows the
  cursor with snapped live state writes while the button is held (thumb shown in the accent
  color), `OnChange` fires once at release if the value moved, and the release click Excel
  then delivers is swallowed. A plain tap sets the value at the press point. Loop-free by
  construction: one key-state poll and at most one cursor read per 16 ms frame, and every
  other pump duty keeps running mid-drag. A visible slider on the active sheet keeps the
  pump armed so presses are never missed - the idle cost is that one poll per frame. While
  a drag or keyboard focus moves the value, it shows in a bubble over the thumb.
  Coordinates come from `GetCursorPos` plus a DPI-and-zoom-aware inversion of
  `PointsToScreenPixels`; frozen-pane splits skew that calibration, so keep app surfaces
  unsplit. Floating chrome claims the points it covers, so a press on an open drop list
  or a modal overlay never reaches a track painted underneath. Use `Stepper` for
  precision.
- `SelectBox`: a themed face, caret, and option list in place of the native dropdown.
  `Text` is the placeholder shown while nothing is selected. The list windows to
  `ListRows` rows (eight by default): opening scrolls the selection into view, and
  clickable pager rows at the list edges (arrow plus the count beyond that edge) page the
  window, drawn whenever the list is longer than its window. Picking a new item writes the
  `WritesTo` state and fires `OnChange`; re-picking the selected item only closes the
  list, the same rule `RadioGroup` follows for its selected row. `ItemEnabled(position,
  False)` leaves an item showing but out of reach: it reads muted, a click on it does
  nothing, and the keys and type-ahead pass over it; a selection already on it stays, and
  `IsItemEnabled` reads it. `AddGroup "Fruit"` appends a group header, a bold, muted row
  flush left that labels the items after it. A header takes a place in the item list, so
  `ItemCount` counts it and later item numbers count past it, and it is never picked,
  selected by `Value`, or reached by the keys. Both marks follow their items through
  inserts and removals.
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
  panel's right edge, 18-point targets, and move the window a page at a time (the arrow
  with nothing beyond it reads muted), and the header counts stay honest about totals. `Reorderable` adds up and down arrows at the
  right end of the chosen panel's header: they move the chosen side's selected rows one
  place as a block, stopping at the ends, and Alt+Up and Alt+Down do the same from its
  rows (the cursor's row when nothing is selected). The chosen order is part of the
  value, so a move writes `WritesTo` and fires `OnChange`. Typing while the list has
  the keys filters the panel its cursor is in: the header shows the filter in quotes and
  how many rows show (`Available "ap" (2 of 5)`), the keys walk the rows that show,
  Backspace takes a letter off, and Esc clears the filter before it leaves. Moves under
  a filter take only rows that show: `>` leaves selected rows the filter hides where
  they are, still selected, and `>>` moves every row that shows. Space still toggles the
  cursor's row, so a filter cannot hold a space.
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
  edge) page the window for the mouse, drawn whenever the matches outrun the window.
  When the text names an item, as it does after a pick, the list reopens unfiltered and
  scrolled to that item, and Down walks on from it; the first edit filters again. While
  the text filters the list, each row bolds what the text matched, and the first item
  that begins with the text shows the rest of its name after the caret in muted ink.
  Right at the end of the text, or Tab, takes it in the item's spelling; Enter commits
  what was typed. A click on a focused combo's text places the caret, and the arrow at
  its right edge toggles the list. `RestrictToItems` limits a float combo's commits to
  its items (see [Field rules](#field-rules)).
- `DatePicker`: a date field. Its face shows the date in the system's short date, or in
  a `Format$` pattern given with `DateFormat "yyyy-mm-dd"`, and `Text` is the placeholder
  it shows in muted ink while it holds none. A click, or Alt+Down, F4, Space, or Down
  while it has the keys, opens a month calendar under the face, over it when there is no
  room below: the month between two arrows, the weekday initials from the system's
  first day of the week, and six weeks of days. The date held fills with the accent,
  today wears an accent ring, and days of other months read muted. A day picks: the
  calendar closes, and a new date writes to `WritesTo` as a VBA `Date` and fires
  `OnChange`. `DateRange earliest, latest` limits what picks, with either side left
  out; days outside read muted and take no clicks, and the arrow and page keys stop at the
  range's first and last day, where a calendar opened with no date starts when today lies
  outside it. `PickDate` sets the date from code
  and `Value 0` clears it, both writing nothing and firing nothing, and `PickedDate`
  reads it as a `Date`, or `Empty` while there is none. `BindValue` takes a `Date` or a
  date serial. The calendar is the control's list: one list is open per app, and a
  press off the calendar and face, a click on another control, Esc, or a move of the
  grid selection closes it.
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
  fires once per toggle, select-all included. Typing while the list has the keys filters
  it: rows the filter excludes hide and the rest close up, the select-all caption shows
  the filter and counts the rows that show (`Select all "gr" (1/2)`), select-all checks
  or unchecks only those rows, Backspace takes a letter off, and Esc clears the filter
  before it leaves. Hidden rows keep their checks, and `WritesTo` still carries every
  checked item. Without the select-all header nothing shows the filter.
- `Image`: a picture as a control - a rounded rectangle whose fill is the picture, so it
  clicks, adopts, and snaps back like everything else, and the picture embeds in the
  workbook. `Source(path)` takes a file path (no URLs) and loads once per distinct path;
  `BindSource(key)` drives it from state. The image stretches to the declared rectangle;
  a missing source renders a themed placeholder, but an already-embedded picture is kept
  even after its source file goes away.
- `Toggle`: the pill switch for booleans. Switched on, the knob takes the theme's
  `OnPrimary` ink, as Windows draws it, so it stays visible on every accent track. A flip,
  by click, key, or bound state, glides the knob across in about 200 ms.
- `MenuButton`: a button that drops a menu of commands. `AddCommand "Export",
  "Module.Export", "Download"` adds a row with its text, the Public procedure it runs, and
  an optional icon, which shows in the row's gutter. `AddGroup` labels a section, and
  `ItemEnabled n, False` leaves a command showing but out of reach. A pick closes the menu
  and runs the command's handler with the menu as `ReDimUI.Sender`, whose `LastCommand`
  names the command; a command without a handler runs the menu's `OnClick` instead. The
  face keeps the menu's own `Text` and looks like a button in its variant, `Secondary` by
  default, with a caret. The menu runs as wide as its widest command needs, never
  narrower than the button, and ends at the button's right edge when it would pass the
  window's. It shares the drop list's windowing, keys, and dismissal, and `ClearItems`
  drops the commands.
- `Expander`: a collapsible section. Its header shows a chevron and its `Text` in bold;
  a click, or Space and Enter while it has the keys, opens and closes it, and Right opens
  and Left closes. Controls join its panel with `InExpander "adv"`, the same as
  `OnTab "adv", 1`, and show only while it is open, under the same rules as a tab's
  panel. It starts closed; `Expanded True` opens it from code, and `IsExpanded` reads it.
  A user's toggle writes `True` or `False` to `WritesTo` and fires `OnChange`. In a
  `Stack`, what follows the panel moves down when it opens and back up when it closes.
- `Badge`: a small pill holding a count or a status word, such as "3" or "Overdue". It
  takes the primary color unless a variant says otherwise (`Success`, `Warning`,
  `Danger`, `Secondary`), and widens to its text: the rectangle's left and top place it,
  its height sets the pill's, and its width is a minimum, so `AtRect(24, 24, 0, 18)`
  fits the text exactly.
- `Sparkline`: a small trend line. `ValuesFrom` takes its numbers, in order, from an
  array, a Collection, a Range (row by row), or a ROneCOne sequence, leaving out blanks
  and text. The line runs from the rectangle's left to its right with its lowest value at
  the bottom and its highest at the top, and a dot marks the last value; equal values run
  along the middle, and one value is a dot alone. It takes the primary color unless a
  variant says otherwise, and screen readers hear a summary: "Trend, 12 values, low 380,
  high 820, last 820".

```vba
ui.Sparkline("spend").AtRect(120, 10, 120, 30).ValuesFrom Range("Spend!B2:B13")
ui.Sparkline("errors").AtRect(120, 50, 120, 30).Danger.ValuesFrom Array(2, 8, 3, 12, 4)
```

- `Skeleton`: a loading placeholder in the muted color. On its own it is one rounded
  block; `SkeletonLines 3` draws the bars of three text lines instead, the last at 60
  percent of the width. It pulses toward the surface color and back every 1.4 seconds,
  and holds still when motion is reduced. Give it the rectangle the content will take,
  and hide it with `Visible False` when the data arrives.
- `Tabs`: a tab strip, one tab per item from `Items`, `ItemsFrom`, and the item APIs.
  Each tab is as wide as its text; when the texts need more than the strip's width the
  tabs share it equally and a text that does not fit ends in an ellipsis. The tab shown
  is bold over an accent bar that glides to the tab shown, the first by default;
  `Value(n)` shows tab n and
  `CurrentValue` reads it. A click, or Left and Right (wrapping) and Home and End while
  it has the keys, shows a tab, writes its text to `WritesTo` (its value, when it has
  one), and fires `OnChange`. A programmatic `Value` writes and fires nothing.

  Controls join a tab's panel with `OnTab`: `ui.CheckList("feat").OnTab "tabs", 2`
  shows the list only while tab 2 is shown. A panel is any set of controls in the
  same app, placed under the strip as usual. A control's own `Visible` and
  `BindVisible` still decide whether it shows on its tab, and one set while its tab
  is hidden waits for the tab. A panel that hides closes its open list, and a field
  on it that has keyboard focus commits and lets focus go. Hiding the `Tabs` control
  hides every panel, `Remove` shows them all, and a `Tabs` control can itself sit on
  another's panel.

```vba
ui.Tabs("tabs").AtRect(24, 24, 360, 32).Items("General", "Advanced").WritesTo "tab"
ui.TextInput("name").AtRect(24, 70, 200, 22).OnTab "tabs", 1
ui.TickBox("beta").AtRect(24, 70, 200, 18).Text("Beta features").OnTab "tabs", 2
```

- `Table`: a data table. `Columns("Item", "Qty")` names the columns and `AddRow "Pear", 3`
  appends a row; `TableFrom source` replaces the rows from a Range or a two-dimensional
  array, its first row naming the columns unless `firstRowHeads:=False`. `ClearRows`
  empties it. Columns share the width by the length of their content unless
  `ColumnWidths` gives points (0 keeps a column's share), and `ColumnFormat 3, "0.00"`
  shows a column through a `Format$` pattern. A column whose filled cells are all numbers
  or dates aligns right, header included; text aligns left and cells too long for their
  column end in an ellipsis. A click on a header sorts by that column, ascending and then
  descending, with an arrow on the header; the sort is stable, numbers and dates sort by
  value ahead of text, text sorts without case, and empty cells sort last either way.
  `SortBy n, descending` sorts from code and `SortBy 0` restores the order the rows came
  in; sorting writes nothing and fires nothing. A click on a row selects it: it fills
  with the accent, its first cell writes to `WritesTo`, and `OnChange` fires. `Value(n)`
  and `CurrentValue` set and read the selection as the row's number in the order the
  rows came in, whatever the sort. Rows show as many as the height holds; more add a
  footer that counts them (`6-10 of 14`) with arrows that page. `RowCount`,
  `CellValue(row, column)`, `SortColumn`, and `SortedDescending` read it back.

  Typing while the table has the keys filters it: a row shows when any cell, as shown,
  holds the text in any case, and the footer names the filter (`Filter "ap": 1-2 of 2`).
  Backspace takes a letter off, Esc clears the filter before it ends focus, and
  `FilterRows "ap"` filters from code; `ShownRowCount` counts the rows it lets through.
  Ctrl+C copies the selected row, or every row shown when none is selected or the filter
  hides it, under the header row as tab-separated text that pastes into cells, dates as
  yyyy-mm-dd so the sheet reads them back as dates. `ExportTo Range("H1")` writes the header and the rows
  shown, filtered and sorted, from that cell down in one write.
  `EmptyText "No orders yet"` words the row an empty table shows, and a filter that
  matches nothing reads "No rows match". `OnRowOpen "Module.Proc"` runs on a double
  click on a row, or on Enter while the table has the keys and shows its selected row,
  with the table as `ReDimUI.Sender` and the row in its `CurrentValue`.

```vba
ui.Table("orders").AtRect(24, 24, 360, 160).TableFrom Range("Orders!A1:C40")
ui.Table("orders").ColumnFormat 3, "#,##0.00"
ui.Table("orders").WritesTo "orderId"
ui.Table("orders").EmptyText("No orders yet").OnRowOpen "Orders.OpenOrder"
```

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
- The numeric keypad types into a focused field as the top row does, its decimal key
  typing the decimal separator Excel uses.
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
- `Required` marks the field: an empty commit, or one of only spaces, shows "Required"
  (the `Required` [UI text](#ui-text)) as its validation message, the same way
  `Validates` shows one, and the check function never sees the empty text.
  `Required emptyMessage:="Enter an email"` words it; `Required False` lifts it.
- `ErrorText "That name is taken"` shows an error from outside the field, a server's
  answer for one, with the same danger border and message line. It stays until
  `ErrorText ""` clears it. A validation message shows in its place while there is one.
- `Masked` shows a float TextInput's characters as dots, focused or not, as a password
  box does. Copy and cut take nothing from it, while `InputValue` and `WritesTo` keep
  the text. `Masked False` lifts it.

A save button checks the whole form at once with `ui.ValidateAll`, which runs every field's
`Required` and `Validates` check, shows each message, and puts focus on the first field
that failed, turning to its tab first:

```vba
Public Sub SaveClicked()
    If Not ReDimUI.SenderApp.ValidateAll Then Exit Sub
    ' every field passed: save
End Sub
```

Any control, not only a field, takes the two adornments. `Caption "Email"` sets a label in
text a point smaller than the control's, above its rectangle, so leave room for it there;
a `Required` control's caption ends in an asterisk in the danger color. `Hint "We never
share it"` sets helper text in muted ink below the control, where a float field's message
takes its place while one shows and a `MaxLength` count keeps the right edge. Both follow
the control's visibility, a disabled control's caption reads muted, and a click on the
caption acts as a click on the control.

```vba
ui.TextInput("email").AtRect(24, 24, 220, 22).Placeholder "name@example.com"
ui.TextInput("email").Caption("Email").Hint("We never share it").Required
ui.TextInput("email").Validates "Checks.Email"
ui.TextInput("qty").AtRect(24, 70, 80, 22).Numeric allowDecimal:=False
ui.TextInput("find").AtRect(24, 116, 220, 22).OnInput "Search.Refilter"
ui.TextInput("find").DebounceMs 250
ui.ComboBox("fruit").AtRect(24, 162, 220, 22).Items("Apple", "Banana").RestrictToItems
```

## Layout

A `Stack` places the controls that join it with `InStack`, so a form needs no
coordinates past the stack's own. Members flow down it in the order they joined, `Gap`
points apart (8 by default), from its top-left corner inside `Padding` (0 by default).
The flow counts what a control carries outside its rectangle: the row a `Caption` takes
above it, and the note row under a field that keeps one, which is any field with a
`Hint`, `Validates`, `Required`, `ErrorText`, or `MaxLength`. The row stays reserved while
no message shows, so a message that comes and goes moves nothing.

- A member keeps its own size, from `Sized(w, h)` or its defaults; its position is the
  stack's from then on. `InStack ""` takes it out where it stands.
- A hidden member gives its place to the next, whether `Visible False`, a
  `BindVisible`, a hidden tab, or a closed `Expander` hid it. When a member grows,
  shrinks, gains a caption, or hides, the members after it move once the change has
  drawn. Only members whose place changed are redrawn.
- `Stretch` gives the controls the stack's width, less padding.
- `Across` runs the members left to right, their tops level below the tallest caption
  among them.
- The stack's height follows its content, and so does its width when it runs across or
  was given a width of 0.
- A stack can sit in another, such as a row of buttons across at the foot of a form.
- A stack is clear and takes no clicks; `Fill(color)` paints it as a panel. It is drawn
  behind its members. `Visible False` hides it with every member, and a member's own
  `Visible` still applies inside.

```vba
ui.Stack("form").AtRect(24, 24, 280, 0).Gap(10).Stretch
ui.TextInput("name").Sized(100, 22).Caption("Name").Required.InStack "form"
ui.TextInput("email").Sized(100, 22).Caption("Email").Hint("We never share it").InStack "form"
ui.Expander("more").Sized(280, 28).Text("More options").InStack "form"
ui.TickBox("news").Sized(200, 18).Text("Newsletter").InExpander("more").InStack "form"
ui.Stack("buttons").Across.Gap(8).InStack "form"
ui.Button("ok").Sized(90, 30).Text("Save").Primary.InStack "buttons"
ui.Button("cancel").Sized(90, 30).Text("Cancel").Secondary.InStack "buttons"
```

## Command palette

`ui.CommandPalette` gives an app a searchable list of everything it can do, the way code
editors do. Ctrl+Shift+P opens it while the app's sheet is in front, and so does
`ui.OpenCommandPalette` from code. It opens as a field at the top of the visible window,
over a list with the keys in the field. The list holds, in this order:

- the app's own entries, added with `ui.AddCommand "Export to CSV", "Module.Export",
  "Download"`;
- every shown, enabled button with a click handler, under its text;
- every `MenuButton` command that can be picked, as "Menu: Command".

Typing filters the list and bolds what matched; the arrows walk it. Enter or a click
closes the palette, gives the keys back to the control that had them, and runs the
entry. An app entry runs its handler with the app as `ReDimUI.SenderApp`, whose
`LastCommand` names it. A button entry clicks the button, with its debounce, busy state,
and handlers. A menu entry runs through the menu. Esc closes the list and then the
palette, and so does a click anywhere else; leaving it never runs what was typed. A
text listed twice keeps its first entry.

```vba
ui.AddCommand "Clear the form", "Form.ClearAll", "Delete"
ui.CommandPalette
```

## Icons

`Icon("Save")` on a `Button` or `Label` draws a Windows icon before the control's text, two
spaces ahead of it, or alone when the control has no text. The icon draws in
`ReDimUI.IconFont` and the text keeps the theme's font. An icon-only button reads to screen
readers as its icon's name, "Delete, button", unless `AltText` says otherwise. `Icon ""`
takes the icon off.

Names, case aside: Add, Remove, Delete, Edit, Save, Search, Filter, Sort, Settings, Refresh,
Sync, Close, Check, More, Menu, ChevronDown, ChevronUp, ChevronLeft, ChevronRight, ArrowUp,
ArrowDown, Back, Forward, Home, Calendar, Clock, Info, Warning, Error, Help, Download,
Upload, Share, Send, Mail, Phone, Person, People, Lock, Unlock, Star, StarFill, Heart,
HeartFill, Flag, Tag, Pin, View, Copy, Paste, Cut, Undo, Redo, Play, Pause, Stop, Folder,
Document, Print, Attach, Link, Globe, Cloud, Camera, Library, Shop, ZoomIn, ZoomOut, List,
Keyboard, and Power. `ReDimUI.IconNames` returns the same list. Any other character of the
icon font works too, passed as one character: `Icon(ChrW(&HE8B7&))`.
`ReDimUI.IconGlyph("Mail")` returns a named icon's character for text of your own, which
needs the icon font on that character.

```vba
ui.Button("save").AtRect(24, 24, 110, 30).Icon("Save").Text "Save"
ui.Button("trash").AtRect(140, 24, 36, 30).Icon("Delete").Danger
ui.Button("inbox").AtRect(24, 70, 110, 30).Icon("Mail").Text("Inbox").BadgeText "3"
```

## Pointer

Shape macros run when the mouse button comes up, so what a pointer does before that
(resting, pressing, holding, dragging) is watched by the pump's frames. The watches read
the pointer at most once a frame and act only while their sheet is in front.

- Hover and press looks, with `ui.PointerEffects`: the fill under the pointer moves
  toward its ink, 8 percent on hover and 16 while the left button that went down on it
  stays down, the state layers Material Design draws. A press that went down elsewhere
  presses nothing it crosses. Buttons, toggles, select, date, and field faces, stepper
  buttons, a slider's thumb, a transfer list's rows, move buttons, and paging arrows, a
  tab, a calendar's days and month arrows, and a table's headers, rows, and paging arrows
  tint; a check box's or radio button's edge takes the accent; an open drop list's highlight
  follows the pointer as it moves, as a Windows list's does, and Enter from a combo or a
  keyboard-focused select takes that row. Off by default, since while it is on the pump
  reads the pointer every frame the app's sheet is in front.
- Tooltips: `Tooltip "Saves the form"` shows the note once the pointer has held still on
  the control for the Windows tooltip delay, the double-click time; a pointer that moves
  more than a few points starts the wait over, so passing over a control raises nothing.
  The note is in inverse colors, as wide as its words, and wraps past 240 points. Under a
  control no taller than 48 points it sits clear of the control, below its hint or
  message, or above its caption when the window has no room below. On a taller control,
  such as a table, it sits under the pointer, flipped above or left to stay in view, and
  goes as soon as the pointer reaches it, so a click there lands on the control. It also
  goes when the pointer leaves, and a press puts it away until the pointer leaves the
  control. An open drop list or calendar shows no note over its rows. `DisabledReason
  "Fill in the address first"` is the note while the control is disabled, and a note
  showing when its words change, as when the control is disabled under the pointer,
  takes the new words. Screen readers get the note with the control's alternative text.
  An app with a tooltip reads the pointer every frame its sheet is in front, as
  `PointerEffects` does.
- Hold-to-repeat: a press held on a stepper's minus or plus, a drop list's pager, a
  transfer panel's paging arrow, a table's footer pager, or an open calendar's month
  arrow repeats after the Windows keyboard repeat delay, at the
  keyboard repeat rate, while the pointer stays on it. A held stepper writes its
  `WritesTo` state at each step and fires `OnChange` once at the release, as a slider
  drag does, and the click Excel delivers for the release is swallowed.
- Transfer gestures: a double click on a row, two clicks within the Windows double-click
  time, moves it across; click a selected row again after that time to deselect it. A
  press on a row dragged along its panel selects the rows from the pressed one to the one
  under the pointer. Dragged onto the other panel, it outlines that panel in the accent,
  and the release there moves the pressed row, or the whole selection when the row is
  part of it. `OnChange` fires once per move.
- The pump: a visible stepper or transfer list on the active sheet keeps the pump armed
  for its press watch, as a slider does, at the cost of a key-state poll a frame. A
  sheet coming back to the front arms the pump again for the watches it holds.

## Keyboard focus for every control

Every interactive control takes keyboard focus: buttons, toggles, tick boxes, radio groups,
steppers, sliders, selects, check lists, transfer lists, tab strips, date pickers, tables,
images with a click handler, and float fields. Focus comes from the keyboard or from code: Tab and Shift+Tab walk the app's
controls, `component.Focus` and `ui.FocusFirst` place it, and a modal takes it. A click
focuses only a text field, since someone who clicks a button in Excel expects the grid to
keep the keys.

- Tab order: controls with a positive `TabIndex` come first, in ascending order, then the
  rest in creation order. `TabIndex(-1)` leaves a control out of Tab while `Focus` still
  reaches it. Tab wraps within the app. Focus that lands outside the visible window
  scrolls the control into view by rows and columns, leaving the selection where it is.
- `ui.FocusRings` rings a focused control in the accent, three points outside its bounds;
  `FocusRings False` turns the ring back off. It is off by default: focus still moves and
  takes keys, and nothing marks the control. A text field shows focus on its own border
  either way, and a check list or transfer list shows a dotted cursor on the row its
  arrow keys have reached.
- Focus ends on Esc, on a click on another control, on a press anywhere off the control
  once the button comes back up, or when the grid selection moves: the watch that ends a
  text field's focus, for every kind. A control that turns hidden or disabled gives focus
  up.
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
- Shortcuts: `Shortcut "^s"` clicks the control on Ctrl+S while its sheet is in front,
  exactly as a mouse click does: the same handler, the control as `ReDimUI.Sender`, and
  nothing while it is disabled, busy, or hidden, when the next control declaring the key
  takes it instead; while a modal is up, only its buttons do. It goes on a control a
  click acts on as a whole: a Button, Toggle, TickBox, Expander, SelectBox, MenuButton,
  ComboBox, DatePicker, float TextInput, or an Image, Label, or Card with a click handler.
  Any other control raises an error, since a slider or a list has no one click to give.
  The code is an `Application.OnKey` code with Ctrl (`^`) or Alt (`%`), Shift (`+`)
  optional, such as `"^+e"` for Ctrl+Shift+E, or a function key from `"{F1}"` to
  `"{F15}"`. The key is a character, `~` for Enter, or a key name OnKey knows, in braces,
  such as `"^{DEL}"`. A key a focused field types is refused, and so is Ctrl+Alt with a
  character: Windows sends AltGr as Ctrl+Alt, and many keyboards type characters with it.
  `""` removes the shortcut. A focused field keeps its own editing chords (Ctrl+A,
  Ctrl+C, and the rest) while it has the keys, and the shortcut takes its key back when
  the field lets go; any other focused control lets a shortcut's key through, so Ctrl+Z
  reaches an Undo button while a list has focus. The tooltip shows the shortcut, as
  "Saves the form (Ctrl+S)" or alone, and the alternative text names it after the
  control's role: "Save, button, Ctrl+S, Saves the form". Like access keys, shortcuts
  bind only while their sheet is active and take the key from Excel while bound, and
  removing the control gives its key back. Give each key one owner: a key that is also
  the app's `HotKey` goes to whichever bound it last, and releasing either releases it.
  Since the tooltip names it, a shortcut keeps the pump reading the pointer while its
  sheet is in front, as a tooltip does. ReDim lists the keys it has bound in a hidden
  workbook name, `rdm_bound_keys`, so `Shutdown`, which runs as the workbook closes,
  releases them even after a reset of the VBA project lost ReDim's own record; a key left
  bound by such a reset clicks nothing and goes back to Excel on its first press.

| Control | Keys |
|---|---|
| Button, Image | Space or Enter clicks. |
| Toggle, TickBox | Space toggles. |
| RadioGroup | Arrows move the selection, wrapping at the ends; Home and End jump. Each move fires `OnChange`. |
| Tabs | Left and Right show the tab beside, wrapping at the ends; Home and End jump. Each move fires `OnChange`. |
| Expander | Space or Enter opens or closes it; Right opens and Left closes. Each change fires `OnChange`. |
| MenuButton | Closed: Space, Enter, Down, Up, Alt+Down, or F4 opens the menu with the first command highlighted. Open: the arrows, Page Up, Page Down, Home, and End move the highlight past disabled commands and headers, letters jump by name, Enter or Space runs the highlighted command, and Esc, F4, Alt+Up, or Tab closes. |
| Table | Up and Down move the selection a row in the order shown, Page Up and Page Down a page, Home and End to the first and last row; the rows scroll to keep it in view, and each move fires `OnChange`. Other characters filter the rows, Backspace takes one off, and Esc clears the filter. Ctrl+C copies the selected row or the rows shown, and Enter opens the selected row when the table has `OnRowOpen`. |
| DatePicker | Closed: Alt+Down, F4, Space, or Down opens the calendar on the date held, or today. Open: Left and Right move a day, Up and Down a week, Page Up and Page Down a month, Home and End to the month's first and last day; Enter or Space picks the day reached, and Esc, F4, or Alt+Up closes. A dashed ring marks the day reached. |
| Stepper | Up and Right step up, Down and Left step down, Page Up and Page Down step ten times, Home and End jump to the range ends. |
| SlideBar | Arrows move a step, Page Up and Page Down a tenth of the range in whole steps, Home and End go to the ends. |
| SelectBox | Closed: arrows, Home, End, Page Up, and Page Down change the selection, and Space, Alt+Down, or F4 opens the list. Open: they move the highlight; Enter, Space, or Alt+Up takes it, Tab takes it and moves on, and Esc or F4 closes. Letters jump to the next item that starts with them, open or closed: letters typed within a second build a prefix, and one letter typed again steps through its items. Every move passes over disabled items and group headers. |
| CheckList | Up and Down move the row cursor, the select-all header included; Home and End jump; Space toggles the cursor's row; Ctrl+A checks every row that shows. Other characters filter the list, Backspace takes one off, and Esc clears the filter. |
| TransferList | Up and Down move the row cursor, Left and Right switch panels, Space toggles the cursor's row in the selection, and Enter moves the panel's selection across, or the cursor's row when nothing is selected; Ctrl+A selects every row of the cursor's panel that shows. Other characters filter the cursor's panel, Backspace takes one off, and Esc clears the filter. With `Reorderable`, Alt+Up and Alt+Down move the chosen side's selection. |
| TextInput, ComboBox | The editing keys under [Text editing](#text-editing); a combo also opens with Alt+Down or F4, closes with Alt+Up, pages its list with Page Up and Page Down, and takes its suggestion with Right at the end of the text or Tab. |

## Accessibility

- Alternative text: ReDim writes a description on each control's shape from its kind, text,
  and state, such as "Save, button", "Agree, checkbox, checked", "Drop-down, South",
  "Tabs, Advanced selected, tab 2 of 3", "Progress, 40 percent" (in 5 percent steps),
  with ", unavailable" when disabled. A `Masked` field reads "Edit field, masked" and
  never its text. Labels, cards, and toasts carry none, since their text is what a
  reader announces. A `Tooltip`
  follows the description, or the `DisabledReason` while the control is disabled.
  `AltText` replaces the description.
- Keyboard: every control works without a mouse, and focus shows as a ring or a field
  border (see [Keyboard focus for every control](#keyboard-focus-for-every-control)).
- State without color: a drop list's current item and a transfer panel's selected rows
  show a check, not only a fill.
- Contrast: `ThemeHighContrast` passes WCAG AA for every pairing the controls draw, and
  `theme.ContrastReport` checks any theme, a custom one included. The light and dark
  presets keep their look, and their reports show one shortfall each: the field edge
  (`Border` on `Surface`) reaches 1.32:1 and 2.01:1 against the 3:1 non-text minimum.
- Motion: toasts stop sliding and fading, a switch's knob and a tab strip's bar land
  without gliding, and a `Skeleton` stops pulsing, when Windows animations are off or
  `ReduceMotion True` is set. The spinner keeps turning; it is status, not decoration.
- Targets: transfer paging arrows and the toast close button are 18 points square, the
  WCAG 2.5.8 minimum of 24 pixels at 96 DPI.

## UI text

ReDim draws and announces a few words of its own, such as a table's "No rows" or the
", button" a screen reader hears. `ReDimUI.SetUIText key, words` replaces one for an app in
another language; `ReDimUI.UIText(key)` reads it, `ReDimUI.UITextKeys` lists every key
for a translation table, and `ReDimUI.ResetUIText` restores English. Keys match in any
case, and an unknown key raises. `{0}`, `{1}`, and on stand for the numbers and names of
the moment, in the order the table gives, and may move within the words. Set the words
before building; a control already drawn takes new ones as it repaints, and `Render`
repaints an app at once. Month and weekday names come from Windows already.

```vba
ReDimUI.SetUIText "NoRows", "Keine Zeilen"
ReDimUI.SetUIText "RowRange", "{0}-{1} von {2}"
ReDimUI.SetUIText "AltButton", "{0}, Schaltflaeche"
```

| Key | English | Key | English |
|---|---|---|---|
| `NoRows` | No rows | `AltButton` | {0}, button |
| `NoRowsMatch` | No rows match | `AltBusy` | busy |
| `NoItems` | No items | `AltProgress` | Progress, {0} percent |
| `NoMatches` | No matches | `AltSpinner` | Busy |
| `RowRange` | {0}-{1} of {2} | `AltSwitchOn` / `AltSwitchOff` | Switch, on / Switch, off |
| `RowRangeEmpty` | 0 of 0 | `AltEditField` | Edit field |
| `FilterNote` | Filter "{0}": {1} | `AltMasked` | masked |
| `SelectAll` | Select all ({0}/{1}) | `AltComboBox` | Combo box |
| `SelectAllFiltered` | Select all "{0}" ({1}/{2}) | `AltMenuButton` | {0}, menu button |
| `Available` | Available | `AltDropDown` | Drop-down, {0} |
| `Selected` | Selected | `AltChecked` / `AltNotChecked` | {0}, checkbox, checked / not checked |
| `PanelFiltered` | {0} "{1}" ({2} of {3}) | `AltRadioGroup` | Radio group |
| `MoreItems` | {0} more | `AltItemSelected` | {0} selected |
| `TypeCommand` | Type a command | `AltStepper` / `AltSlider` | Stepper, {0} / Slider, {0} |
| `Working` | Working... | `AltTransferList` | Transfer list |
| `Required` | Required | `AltCheckList` | Checklist, {0} of {1} checked |
| `OK` | OK | `AltImage` / `AltLoading` | Image / Loading |
| `Cancel` | Cancel | `AltExpanded` / `AltCollapsed` | {0}, expander, expanded / collapsed |
| `Column` | Column {0} | `AltTable` / `AltTableOneRow` | Table, {0} rows / Table, 1 row |
| `ImageNotFound` | (image not found) | `AltSortedAscending` / `AltSortedDescending` | sorted by {0}, ascending / descending |
| | | `AltShown` | {0} shown |
| | | `AltDatePicker` | Date picker, {0} |
| | | `AltTabs` / `AltTabSelected` | Tabs / {0} selected, tab {1} of {2} |
| | | `AltUnavailable` | unavailable |
| | | `AltSparkline` | Trend, {0} values, low {1}, high {2}, last {3} |
| | | `AltSparklineEmpty` | Trend, no values |

`Working` is what a busy button shows unless `BusyText` words it, `Required` is the message
unless `Required` is given one, and `OK` and `Cancel` label `Confirm`'s buttons unless it is
given words. `FilterNote` wraps a table's `RowRange` while a filter holds, `SelectAll` heads
a check list, `Available`, `Selected`, and `PanelFiltered` head a transfer list's panels,
`MoreItems` is a drop list's pager row, and `TypeCommand` is the command palette's hint.

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
