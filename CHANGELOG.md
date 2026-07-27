# Changelog

## 0.14.1 - 2026-07-26

- The gallery's generated demo image was a solid-color PNG, which is
  indistinguishable from a plain filled shape and read as "nothing
  showing". The Image control was loading it correctly; the fixture
  could not prove it. The demo logo is now composed from overlapping
  shapes and text on a chart canvas before export, so the picture is
  unmistakable.

## 0.14.0 - 2026-07-26

- `Image`: a picture control drawn the family way - a rounded rectangle
  whose fill is the picture, so corners, geometry, adoption, snap-back,
  and click dispatch all behave like every other control, and the
  picture embeds in the workbook. `Source` takes a file path (no URLs -
  the honest VBA primitive) and loads it once per distinct path;
  `BindSource` drives the source from a state key like the other
  bindings. The image stretches to the declared rectangle. A missing or
  unloadable source renders a themed placeholder, except when the shape
  already carries an embedded picture: a picture saved with the
  workbook outlives its source file by design. The gallery shows a
  self-generated logo (chart-area export) beside the checkbox list.

## 0.13.0 - 2026-07-26

- BREAKING: the persistence layer is removed entirely. `Persist` and
  `ClearPersisted` are gone, and hidden `rdm_s_*` workbook names are no
  longer written or read. The state store is deliberately in-memory and
  session-scoped; durability belongs to the host application, which can
  read the store with `State`/`StateOrDefault`, save it wherever fits (a
  hidden sheet, workbook names, a file - ROneCOne's JSON serializer is
  on board), and reseed on build with `SetStateDefault`, which never
  clobbers a value already in play. The Navigator demo's settings are
  session-scoped accordingly.

## 0.12.0 - 2026-07-26

- `CheckList`: a drawn checkbox list. One box-and-caption row per item,
  any number checked, and a select-all header row on by default
  (`WithSelectAll False` to opt out) following the standard tri-state
  conventions: the header box reads empty, checked, or dashed for a
  mixed state, clicking it checks everything unless everything is
  already checked, and the caption shows the live count. Box and
  caption both toggle their row, so the whole row is a hit target.
  `ItemsFrom` and the item APIs feed the rows; `CheckedFrom` seeds the
  checked set by text; `SetItemChecked`, `IsItemChecked`, and
  `CheckedCount` are the programmatic surface, silent by the usual
  contract. Checks follow their items through inserts and removals.
  `WritesTo` carries the checked items joined with ", "; `OnChange`
  fires once per toggle, bulk select-all included. The gallery gains
  an options list beside the crew roster.

## 0.11.1 - 2026-07-26

- TransferList rows multi-select. Ctrl+click cannot work on a drawn
  control: Excel reserves Ctrl+click (and Shift+click) on a
  macro-assigned shape for selecting the shape itself, so the macro
  never runs - the same platform rule the framework documents as the
  design-time escape hatch. Plain clicks toggle rows in and out of the
  selection set instead, both panels independently; the single move
  buttons transfer every selected row in list order with one state
  write and one OnChange. The single-item flow is unchanged: click
  one, move it.

## 0.11.0 - 2026-07-26

- `TransferList`: a drawn dual listbox (transfer list). Two panels with
  counted headers, selectable rows, and four move buttons between them -
  move one right or left, move all right or left, matching the classic
  control. `Items`/`ItemsFrom` and the item APIs feed the available side;
  `ChosenFrom` seeds the chosen side; `Captions` names the headers;
  `ChosenCount`/`ChosenTextAt` read the result. `WritesTo` carries the
  chosen items joined with a comma and space, and `OnChange` fires once
  per user transfer (row selection fires nothing). Rows render up to the
  panel's capacity and the header counts keep overflow honest. The
  gallery gains a crew roster between the fields and the slider block.

## 0.10.2 - 2026-07-26

- Fix the root cause of new toasts landing in occupied positions:
  CompactToastSlots read `ToastSlot` through `Collection.Item(...)`
  directly, and Friend members are invisible to late binding, so with
  two or more survivors the comparison loop raised error 438 and the
  swallowed error killed the whole compaction - survivors kept their
  slots, the freed low slot went to the next newcomer. The failure
  needed 2+ survivors, which no prior test ever staged; the new
  mid-glide scenario staged it and caught the crash. Candidates now go
  through a typed local (the codebase carries no other late-bound
  member chains).
- Toast choreography tightened around the remaining visual races:
  - A newcomer spawned during a compaction glide entered a slot that was
    model-correct but still visually occupied by a survivor easing
    upward. Newcomers now enter from below the column as drawn - one
    full pitch under the lowest live toast when that is deeper than the
    default entrance hop - so overlap is impossible at any frame.
  - Expired and dismissed toasts fade out over 180 ms before removal and
    compaction, instead of popping out of existence mid-scene.
  - The toast shape is created at its entrance position rather than its
    settled slot, removing a one-paint flash at the target position.

## 0.10.1 - 2026-07-26

- Fix a new toast sometimes landing on top of live ones (reported as a
  fresh toast in slot one while others were on screen). Slots were always
  consistent; the tray ORIGIN was not - every spawn recomputed the
  viewport-clamped rail, so a scroll, zoom, or content-bounds shift
  between spawns moved the rail mid-stream and the newcomer's slot was
  measured from a different origin than the survivors'. The rail is now
  sticky while any toast is alive: the toast that opens the tray fixes
  the origin, newcomers join that column, and the next toast after the
  tray empties re-establishes a fresh clamped origin. Regression test
  scrolls between spawns and asserts the newcomer joins the live column
  exactly one slot pitch below.

## 0.10.0 - 2026-07-26

- `ProtectSurface` makes locked canvas cells unselectable by default
  (`EnableSelection = xlUnlockedCells`): no selection rectangle on the app
  surface, no protected-cell warnings for stray keys, and deliberately
  unlocked cell-backed TextInput cells stay selectable and editable. Pass
  `allowCellSelection:=True` for the previous behavior. Verified with
  message-level keystrokes that OnKey capture - float-field typing and
  HotKey arrows - fires exactly as before with no selectable cell.
- Since grid clicks on such a surface move no selection, click-away
  commit there is Enter, Tab, or clicking any control; the press-edge
  watch still covers shape and chrome clicks.

## 0.9.4 - 2026-07-26

- Replace 0.9.3's event-driven cell-click commit with a frame-driven
  selection poll, after a field report showed the event approach not
  landing. SheetSelectionChange had two structural weaknesses: it runs
  inside the grid's selection mouse loop, an execution context where the
  framework's shape surgery was never proven, and it goes silent whenever
  a host runs with application events off. The focused field's frames -
  the same ticks that blink the caret, observably alive in the field
  reports - now compare the selection against a snapshot taken at focus;
  any move commits. Works with events disabled, runs in proven tick
  context, and also commits on arrow-key moves and sheet navigation.
- The live test drives the poll with events off, exactly as the harness
  runs, so the covered path is the shipped path.

## 0.9.3 - 2026-07-26

- Fix clicking a cell not committing a focused float field (user repro:
  focused combo kept its ring and open list after a cell click). Every
  polled link verified clean under real timer ticks, real cursor
  positions, and Excel's own RangeFromPoint ground truth on a scrolled
  window - the failure is the press itself: the grid's selection mouse
  loop can hold WM_TIMER until the button is already up, so a cell press
  can be invisible to the pump's press-edge poll (shape presses keep
  frames running, which is why sliders never missed). Cell clicks now
  commit event-driven through Application.SheetSelectionChange - the
  selection change is the click signal, no timer involved. Arrow-key
  selection moves commit the same way. The press-edge watch remains for
  clicks that change no selection (ribbon, title bar, other windows,
  re-clicking the selected cell).
- New test seam ForcePressEdge lets live tests drive the in-tick blur
  path without physical input; the widget suite now covers the
  selection-driven commit through a real SheetSelectionChange event.

## 0.9.2 - 2026-07-26

- Clicking away from a focused float field commits with either mouse button:
  the blur watch now treats a right-button press as clicking away too, so a
  context-menu click cannot leave a field silently holding focus. (Left-click
  away has committed since 0.9.0 via the pump press-edge watch; in 0.9.0 the
  invisible-ink bug merely hid it happening.)
- Tab commits and leaves the field, matching the form convention alongside
  Enter; Esc still reverts.

## 0.9.1 - 2026-07-26

- Fix invisible typing in float fields: a field left on the default variant
  implicitly carried the primary style, whose white ink landed on the field's
  white surface fill - the buffer took every keystroke but the text and the
  insertion bar were unreadable, presenting as "typing does not work". Float
  fields now always read in surface ink, like every surface-filled kind. The
  live test asserts ink and ring colors, not just text content.
- TextInput and ComboBox faces are rounded rectangles matching the rest of
  the control family; shapes adopted from older builds are coerced on first
  render.
- Verified with real (message-level) keystrokes end to end, including on a
  ProtectSurface sheet: OnKey capture receives keys with protection active,
  so protection now stays on during focus. Keys outside the bound typing set
  fall through to the grid, where the protected surface answers with Excel's
  usual notice instead of silently entering cell edit.

## 0.9.0 - 2026-07-26

- Cell-free fields. `TextInput` and `ComboBox` placed with `AtRect`, `Below`, or `RightOf`
  are float fields: text lives in the component buffer and renders on the shape, edited
  through a keyboard focus layer instead of a cell. Click to focus (accent ring, blinking
  insertion bar), type, Enter commits (state write plus OnChange on change), Esc reverts,
  and a press anywhere off the field commits - watched by the same pump frames as slider
  drags. One field holds focus at a time. `At` keeps the cell-backed mode of both controls.
- Live combo filtering, the payoff cells could never give: the framework sees every
  keystroke in a float field, so the drop list re-filters as you type instead of waiting
  for a commit. Float fields also type normally on `ProtectSurface` sheets, which lock out
  cell edit entirely.
- Capture uses `Application.OnKey`, bound only while a field is focused and released on
  blur: letters with Shift capitals, digits, space, minus, period, comma, Backspace, Enter,
  Esc. The editing model is append-and-backspace - OnKey cannot report cursor movement.
  `RdxReleaseKeys` unbinds everything regardless of surviving state.
- Focus ring now always uses the theme primary color; previously a default-variant field
  would have drawn a muted ring.
- The gallery's input row goes cell-free and demonstrates live filtering on the protected
  surface. With that, no shipped control or demo needs a worksheet cell.

## 0.8.0 - 2026-07-26

- `ComboBox`: an editable combo built from existing parts - a cell-backed text face with a
  caret and a drawn, filtered drop list, sharing the item APIs. Excel pauses all VBA during
  cell edit mode, so filtering applies at the moments the platform grants: the caret opens
  the list filtered by the typed text, and an Enter commit auto-suggests partial matches,
  takes an exact match outright, or stays closed as free text. Picking writes the cell, the
  state, and fires OnChange. The gallery gains a fruit combo beside the text input.

## 0.7.3 - 2026-07-26

- Programmatic item APIs for SelectBox and RadioGroup: `AddItem` with optional position,
  `RemoveItem` by index or text, `ClearItems`, `ItemsFrom` accepting a 1D array, a
  Collection, or a Range (one item per non-empty cell), plus `ItemCount` and `ItemTextAt`
  readers. The selected item survives inserts and unrelated removals; removing it clears
  the selection to the placeholder. Shrinking a RadioGroup now sweeps its stale row parts.

## 0.7.2 - 2026-07-26

- Fix drag never engaging on freshly opened workbooks: slider demand kept an armed pump
  alive but nothing armed it after Render, so the press watch was not running until some
  other feature started the pump. Render and Navigate now arm the pump whenever the app has
  pending work. A real-input experiment also confirmed the platform premise: the pump's
  timer keeps firing and COM writes keep succeeding while the mouse button is held down on
  a macro shape, so mid-drag tracking is fully available.

## 0.7.1 - 2026-07-26

- True press-drag sliding. OnAction only fires at mouse up, so 0.7.0's click-then-follow
  mode could not serve the natural press-hold-sweep gesture and actively fought it. The
  pump's frames now watch the left-button press edge themselves, hit-test the cursor against
  the track, and run the drag session: live snapped writes while held, OnChange once at
  release if the value moved, release click swallowed, plain taps still set. A visible
  slider on the active sheet keeps the pump armed so presses are never missed; the idle cost
  is one key-state poll per frame.

## 0.7.0 - 2026-07-26

- Non-blocking sliding: the first click on a SlideBar sets the value and grabs the thumb,
  pump frames follow the cursor with snapped live state writes, and either a second click on
  the slider or a mouse press anywhere else drops it. OnChange fires on the engaging click
  and once on a drop that moved the value. The engaged thumb shows the accent color. One
  cursor read per frame; every other pump duty keeps running while the user slides.
- Cursor pinning becomes opt-in (`PinPumpCursor`). The always-on pin was hiding the hover
  hand on interactive shapes while async work ran; lean frames made the busy-cursor strobe
  it suppressed negligible, so hover affordance now wins by default.

## 0.6.2 - 2026-07-26

- Fix SlideBar click accuracy. PointsToScreenPixelsX is misnamed on current Excel builds:
  probing showed its input is document pixels (slope exactly 1 at any zoom, while scrolling
  shifts the origin by the true pixels-per-point), so treating cursor pixels as points landed
  clicks about a third short. The mapping now derives pixels-per-point from display DPI and
  window zoom, selects the conversion by comparing the measured slope against it (robust to
  either contract), and returns absolute sheet points scroll-proof. A physical regression
  test pins the contract: a fixed screen pixel must map to points shifted by exactly the
  scroll delta, and spans must scale inversely with zoom.

## 0.6.1 - 2026-07-26

- `SlideBar` drops the drag loop and becomes pure click-to-set. The loop violated the
  framework's no-blocking rule, and it could never engage anyway: shape OnAction fires on
  mouse up, so no button is held by the time a handler runs. One cursor read at dispatch,
  snapped to SliderRange, OnChange per click. The GetAsyncKeyState and Sleep declares are
  gone with it.

## 0.6.0 - 2026-07-26

- `SlideBar`: a drawn slider with click-to-set and live drag, correcting 0.4.0's claim that
  this was impossible. GetCursorPos plus a two-point inversion of PointsToScreenPixelsX maps
  the cursor into track fractions at any zoom, and a left-button key-state loop tracks the
  drag. Snaps to SliderRange, writes state live, fires OnChange on release. The pump pauses
  during a drag; the slider repaints from the drag loop itself.
- Gallery: the input frame anchors at C20 (a 15-point row-height arithmetic slip had parked
  it beside the TickBox), and the slider, stepper, and meter now share one state key.

## 0.5.0 - 2026-07-26

- Breaking: the native form control wrappers (`Checkbox`, `Dropdown`, `Slider`) are removed,
  along with the whole form-control rendering path. The drawn family (`Toggle`, `TickBox`,
  `RadioGroup`, `Stepper`, `SelectBox`) covers every interaction with full theme control.
  Migration: `Checkbox` becomes `TickBox`, `Dropdown` becomes `SelectBox`, `Slider` becomes
  `Stepper` (optionally paired with a bound `ProgressBar`).
- The gallery drops its native row; the meter now binds to a drawn Stepper.

## 0.4.0 - 2026-07-26

- Drawn control family completes: `TickBox` (themed checkbox with check glyph and clickable
  caption), `RadioGroup` (single-select rows the native controls never offered), and
  `Stepper` (numeric minus/value/plus honoring SliderRange). With `Toggle` and `SelectBox`,
  every interactive control now has a fully themed shape-drawn form; native wrappers remain
  optional. Pure shapes, no new dependencies.
- The gallery gains a drawn-controls section wired to the state inspector.

## 0.3.1 - 2026-07-26

- `NavBar`: one tab per registered window across the top of a sheet, active tab highlighted,
  refreshed on every navigation, stale tabs pruned when windows unmount. `WindowTitle` gives
  windows display names. The Navigator demo now navigates by tabs.

## 0.3.0 - 2026-07-26

- Window system: `AsWindow` registers an app's sheet as a form-like window; `Navigate` shows
  one window at a time (target first, others very-hidden), `NavigateBack` walks the stack,
  `OnShow` and `OnHide` fire as lifecycle hooks, and `NavigatesTo` turns any component into a
  nav link. Unregistered sheets are never touched, and background windows keep pumping.
- New Navigator demo: three windows with navigation, a back stack, lifecycle counters, and
  persisted settings.
- All demos now ship with `ProtectSurface` applied, using the reopen-safe pattern of
  unprotecting at build start and protecting after Render.

## 0.2.7 - 2026-07-26

- `ProtectSurface`: opt-in UserInterfaceOnly protection for the app sheet, so users cannot
  park the pump in cell edit mode or drag shapes there, while framework writes keep working
  and TextInput cells stay editable. Unmount unprotects. Deliberately scoped to the one
  sheet: ReDim will not disable Excel-wide input or dialogs to protect an animation.

## 0.2.6 - 2026-07-26

- Steadier frames: the multimedia timer resolution rises to 1 ms while the pump is armed and
  is restored on stop, so frames stop quantizing into ~15.6 ms buckets
- Frame deltas and job budgets are measured with QueryPerformanceCounter, because
  GetTickCount64 stays at the coarse system tick on modern Windows and cannot express
  sub-16 ms budgets
- Budget jobs on the default budget yield to the frame rate (8 ms per pass) while anything
  animates; explicit BudgetMs values are always honored

## 0.2.5 - 2026-07-26

- The pump runs animation frames at 16 ms (about 60 fps) while ops, budget jobs, and toast
  expiry keep their 50 ms work cadence, so smoother motion costs no extra task polling.
  Spinner rotation and toast easing are time-based, preserving their speeds at any frame
  rate, and paced jobs are sampled every frame for tighter game timing.

## 0.2.4 - 2026-07-26

- Toast stack compaction: when a toast expires or is dismissed, survivors renumber to slots
  one through N and slide up into the freed positions, and new toasts join below the live
  stack. The tray now moves like a notification tray.

## 0.2.3 - 2026-07-26

- The default toast rail clamps into the visible viewport: the rail position wins whenever it
  is on screen, and visibility wins when the window is narrower than the content or scrolled
  away. Explicit `ToastTray` pins are honored exactly.
- Toasts slide up into their slot on entrance, a pump-driven ease-out of about 300ms.

## 0.2.2 - 2026-07-26

- Toasts sit on a stable rail just outside the content's right edge, with `ToastTray` to pin
  the rail to a range; modal chrome no longer shifts toast placement
- The spinner is a block arc, a ring concentric with its box, so rotation is circular instead
  of the wobbling stroke a plain arc produces; stale arc shapes are rebuilt on adoption
- `BindEnabled` and `BindVisible` take an invert flag; Mission Control's launch button now
  disables while any feed runs

## 0.2.1 - 2026-07-26

- The pump pins `Application.Cursor` while armed and restores it on stop, ending the rapid
  busy-cursor strobe Excel produces when VBA executes at pump frequency
- Every framework shape carries the dispatcher `OnAction`, so plain clicks can no longer
  select and drag modal cards, labels, or other chrome; Ctrl+click remains for design work
- Geometry diffs against the live shape instead of a cache, so manually moved shapes snap
  back to their declared rectangles on the next render
- `FlushDirty` no longer re-assigns `ScreenUpdating` it never changed

## 0.2.0 - 2026-07-25

- SelectBox: a fully drawn themed picker, because Excel fixes the native dropdown's list font
- Checkbox captions render as themed label parts that still toggle the box on click
- State persistence: `Persist`, `SetStateDefault`, and `ClearPersisted` over hidden workbook
  names, so reopened workbooks resume where the user left off
- Relative layout: `Below`, `RightOf`, and `Sized` remove coordinate arithmetic from app code
- Managed hotkeys: `HotKey` and `ClearHotKeys` with automatic release on Unmount and Shutdown
- Render now prunes orphan shapes left by renamed or removed components
- Paced jobs (`PacedMs`) for game loops and animations; Snake and Mission Control use them
- Toasts anchor to the app's content bounds, reclaim freed slots lowest-first, and always use
  surface ink in both themes
- Secondary buttons carry the theme border so they read against surface backgrounds
- The modal overlay covers from the origin to one viewport past the visible area
- Pump ticks no longer toggle ScreenUpdating, ending cursor and scroll bar flicker
- `ReDimUI.Version`

## 0.1.0 - 2026-07-25

- Initial framework: role-tagged `ReDimUI.cls` runtime plus `ReDimHost.bas` host module
- Retained components over worksheet shapes with diffed rendering and idempotent mounts
- State store with text, value, visibility, and enabled bindings
- SetTimer pump stepping ROneCOne tasks without blocking waits; async ops with disable,
  spinner, done, fail, and cancel wiring; budget jobs; cancellation tokens
- Widgets: Button, Label, Card, ProgressBar, Spinner, Toggle, Checkbox, Dropdown, Slider,
  cell-backed TextInput, Toast, shapes-based modal
- Light and dark themes
- Build pipeline over pyOpenVBA, static gate over pyvbaanalysis, live suite over pyvbaharness
  with a real-compiler compile gate
- Demos: Mission Control, Widget Gallery, Snake
