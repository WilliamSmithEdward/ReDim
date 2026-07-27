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
| `ReDimUI.PinPumpCursor enabled` | Opt-in steady arrow cursor while the pump is armed. Off by default so interactive shapes keep their hover hand; turn on if busy-cursor flicker is visible on your hardware. |
| `ReDimUI.AutoPump enabled` | Turn the wall-clock timer off for deterministic runs. |
| `ReDimUI.Shutdown` | Kill the pump and forget every app. Shapes stay. |
| `ReDimUI.ThemeLight`, `ThemeDark` | Theme presets; customize with `WithPrimary`, `WithFont`. |

App ids use letters and digits only. Component ids may add single underscores.

## Windows and navigation

Sheets as forms: an app registered with `AsWindow` becomes a window, and navigation shows one
window at a time among the registered set.

| Member | Purpose |
|---|---|
| `app.AsWindow` | Registers the app's sheet as a window. |
| `ReDimUI.Navigate appId` | Shows the target window's sheet, activates it, then very-hides every other registered window. Sheets that are not windows are never touched. |
| `ReDimUI.NavigateBack` | Pops the back stack; returns False when empty. |
| `ReDimUI.ActiveWindowId` | The currently shown window's app id. |
| `app.OnShow "Module.Proc"` / `app.OnHide "Module.Proc"` | Lifecycle hooks fired after navigation shows or hides the window. |
| `component.NavigatesTo "appId"` | One-declaration nav link: navigates after any OnClick handler. |
| `app.WindowTitle "Home"` | Display name used by navigation chrome; the app id is the fallback. |
| `app.NavBar left, top, tabWidth, tabHeight` | One tab per registered window across the top of this sheet, active tab highlighted. Bars refresh on every Navigate, picking up late registrations and pruning removed windows. |

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
| `SetTheme theme` | Restyle every component. |
| `PrepareCanvas` | Paint the sheet background and hide gridlines. |
| `Toast message, ttlMs` | Transient card on a rail beside the content, clamped into the visible viewport. Toasts slide up on entrance, and when one leaves the survivors slide up to fill its slot; modal chrome never shifts the rail. |
| `ToastTray anchor` | Pins the tray's top-left to a range, exactly and unclamped. |
| `Confirm title, message, okProc, cancelProc, okText, cancelText` | Shapes-based modal. |
| `CloseModal` | Hide the modal set. |
| `Async(opId)` / `CancelAsync opId` / `AsyncError(opId)` | Async ops (see [async.md](async.md)). |
| `Job(jobId)` / `CancelJob jobId` | Chunked or paced background work. |
| `OnError "Module.Proc"` | One-argument sink for swallowed handler failures. |
| `ProtectSurface enabled, allowCellSelection` | Opt-in app-sheet protection (UserInterfaceOnly): users cannot enter cell edit mode or drag shapes there, framework writes keep working, cell-anchored TextInput cells stay editable, and float fields type normally since they never enter cell edit. By default locked canvas cells are also unselectable - no selection rectangle on the app surface, no protected-cell warnings for stray keys - while unlocked TextInput cells stay selectable; pass `allowCellSelection:=True` to keep the whole grid selectable. OnKey capture (fields, HotKey) is unaffected. Protection state does not persist across reopen, so builds should call `ProtectSurface False` first and `ProtectSurface` after Render, as every demo does. Unmount unprotects. |
| `Unmount deleteShapes` | Remove components (and shapes) and forget the app. |

## Component builders

All fluent, all return the component:

- Geometry: `At("B2:D3")` anchors to a range and follows column widths on re-render;
  `AtRect(left, top, width, height)` uses points; `Below(otherId, gap)` and
  `RightOf(otherId, gap)` place relative to another component with `Sized(w, h)` for
  dimensions. Circular relative chains raise a clear error.
- Content: `Text`, `FontSize`, `Bold`, `BusyText`.
- Style: `Primary`, `Secondary`, `Success`, `Danger`, `Fill(color)`, `TextColor(color)`.
- Visibility: `Visible(flag)`, `Enabled(flag)`.
- Values: `Value(number)` (progress, slider, picker index), `Checked(flag)`,
  `SliderRange(min, max, step)`.
- Item lists (`SelectBox`, `RadioGroup`): `Items("A", "B", ...)` replaces;
  `ItemsFrom(source)` replaces from a 1D array, a Collection, or a Range (one item per
  non-empty cell); `AddItem(text, atPosition)` appends or inserts; `RemoveItem(indexOrText)`;
  `ClearItems`; read back with `ItemCount` and `ItemTextAt(position)`. The selected item
  survives inserts and unrelated removals; removing it clears the selection to the
  placeholder. Programmatic mutations re-render but do not write `WritesTo` state or fire
  `OnChange`; those belong to user interaction and explicit `SetState`.
- Bindings: `BindText(key, template)` where `{0}` is the value, `BindValue(key)`,
  `BindVisible(key, invert)`, `BindEnabled(key, invert)`, `WritesTo(key)`. The invert flag
  serves the disable-while-busy pattern: `BindEnabled "anyRunning", True`.
- Behavior: `OnClick "Module.Proc"`, `OnClickAsync "Module.Proc"`, `OnChange "Module.Proc"`.
- Reads: `CurrentValue`, `CurrentText`, `IsChecked`, `IsEnabled`, `IsBusy`, `InputValue`
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
  unsplit. Use `Stepper` for precision.
- `SelectBox`: a themed face, caret, and option list in place of the native dropdown.
- `TransferList`: a dual listbox - two panels with counted headers, selectable rows, and
  four move buttons (`>`, `>>`, `<`, `<<`). `Items`/`ItemsFrom` and the item APIs feed the
  available side, `ChosenFrom` seeds the chosen side, `Captions` names the headers, and
  `ChosenCount`/`ChosenTextAt` read the result. Rows multi-select by click-to-toggle:
  each plain click adds or removes that row from the panel's selection, and `>`/`<` move
  every selected row in list order. (Ctrl+click cannot exist on a drawn control - Excel
  reserves modifier-clicks on macro shapes for selecting the shape itself, the same rule
  behind the design-time escape hatch below.) `WritesTo` carries the chosen items joined
  with a comma and space; `OnChange` fires once per user transfer, and selecting rows
  fires nothing. Rows render up to the panel's height; the header counts stay honest when
  a list overflows what fits.
- `ComboBox`: an editable combo with a caret and a filtered drop list, sharing the item
  APIs. Place it with `AtRect` (or `Below`/`RightOf`) and it is a float field: click to
  focus, type, and the list re-filters on every keystroke. Anchoring to a cell with `At`
  keeps the 0.8.0 cell-backed mode, where Excel's edit-mode VBA pause limits filtering to
  commit moments. In both modes picking an option writes the value, the `WritesTo` state,
  and fires `OnChange`.
- `TextInput`: a text field. Float by default (`AtRect`), cell-backed with `At` when you
  want the value to live in the grid.
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
- `Toggle`: the pill switch for booleans.

## Float fields and keyboard focus

As of 0.9.0 no control needs a cell. `TextInput` and `ComboBox` placed with `AtRect`,
`Below`, or `RightOf` render and edit entirely on their shapes: clicking a field gives it
keyboard focus, an accent ring and a blinking insertion bar appear, and characters go to the
component's text buffer instead of a cell. That sidesteps Excel's edit-mode VBA pause - the
framework sees every keystroke, which is what makes live combo filtering possible - and it
means typing works on `ProtectSurface` sheets, where cell edit is locked out.

Focus mechanics, all automatic:

- One field holds focus at a time; focusing another commits the first.
- Enter and Tab commit: the buffer becomes the value, `WritesTo` state is written and
  `OnChange` fires if the text changed. A combo commit that exactly matches an item takes
  that item. A `MultiLine` TextInput follows textarea conventions instead: Enter inserts
  a newline and keeps focus, while Tab, Ctrl+Enter, and clicking away commit; the value
  carries its newlines into state. Size the rectangle for the lines you expect - roughly
  15 points per line plus 6 points of margin; a line that cannot fully fit is not drawn
  at all, which reads as a missing line.
- Overflow follows the caret. Shape text cannot scroll, so a focused field renders the
  tail window of its buffer - the last lines that fit (multi-line) or the rightmost
  characters that fit (single-line and combo) - with a leading ellipsis marking trimmed
  content. The buffer keeps the complete text and commits intact; unfocused fields show
  their beginning.
- Esc reverts a TextInput to the text it had when focus arrived and fires nothing. On a
  combo, Esc clears instead: the first press empties the value and reopens the full list
  with focus kept; a second press on the empty combo leaves and commits the clear.
- Clicking anywhere off the field commits, through two frame-driven signals. A cell click
  moves the selection, and the focused field's frames poll the selection - the press
  itself can be invisible (the grid's selection mouse loop holds timer messages until the
  button is back up), but the selection it leaves behind is durable state the next frame
  sees. Sheet navigation commits the same way, and none of it depends on application
  events. (Arrow keys no longer move the selection while a field is focused - they edit.) Everything that changes no selection - either mouse button on
  shapes, chrome, or other windows - is caught by the same press-edge watch that drives
  slider drags. The one blind spot: re-clicking the already-selected cell during a starved
  moment changes nothing observable; the next keystroke, click, or frame resolves it. On a
  default `ProtectSurface` sheet the canvas is unselectable, so grid clicks move no
  selection at all - there, commit by Enter, Tab, or clicking any control, which is the
  natural flow on an app surface anyway.
- `InputValue` reads and writes the buffer in float mode, the cell in cell mode.

Capture uses `Application.OnKey`, bound only while a field is focused and released on blur,
so sheet typing is untouched the rest of the time. The bound set is the practical editing
set: letters (with Shift capitals), digits, space, minus, period, comma, Backspace, Del,
the arrow keys, Home, End, Tab, Enter (and Ctrl+Enter), Esc. Editing is full caret
editing: arrows move the insertion point, characters insert at it, Backspace and Del
delete around it, Home and End jump the line edges, and Up and Down move across hard
lines with the column clamped (a long wrapped line counts as one line). While a field is
focused the arrows belong to editing, so they do not move the cell selection. Text
selection (shift-selection, copy/paste) is not modeled. Keys outside the bound set fall
through to the grid as usual; on a `ProtectSurface` sheet Excel answers those with its
protected-cell notice, and protection stays on the whole time - OnKey capture works fine
under protection (verified with message-level keystrokes). Apps that bind arrow HotKeys
should re-arm them after field focus sessions if they mix the two. `ReDimUI.HasKeyboardFocus`
and `ReDimUI.FocusedComponentId` report the current holder; `RdxReleaseKeys` is the panic
release that unbinds everything regardless of state.

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
app.Spinner("busy").AtRect 400, 18, 26, 26        ' right
app.Spinner("busy").AtRect(400, 18, 26, 26)       ' compile error: Syntax error
app.Label("x").At("B2").Text("hello")             ' fine: one argument
```

Mid-chain calls are expression position and may keep their parentheses. The repository's compile
gate (`tests/python/test_compile.py`) catches violations that static analysis cannot.

## Case-insensitive shadowing rule

VBA identifiers are case-insensitive. A local `Dim ownerApp` shadows a property named
`OwnerApp`, so `Set ownerApp = OwnerApp` self-assigns Nothing. Never name a local after a member.
