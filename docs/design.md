# ReDim design

ReDim is a stateful UI framework for Excel worksheets. It renders retained components as worksheet
shapes, binds them to a per-app state store, and drives async behavior through a timer pump
layered over ROneCOne's cooperative task scheduler. No UserForms, no ActiveX, no form controls,
no VBIDE access.

## Naming

The framework is named ReDim. `ReDim` is a reserved VBA keyword, so the code surface is `ReDimUI`
(predeclared class) plus `ReDimHost` (standard module). Shape names use the `rdm_` prefix.

## Runtime shape

Two imported files on top of `ROneCOne.cls`:

| File | Role |
|---|---|
| `src/ReDimUI.cls` | Predeclared, role-tagged class: factory, app, component, async op, job, theme |
| `src/ReDimHost.bas` | OnAction dispatch target, SetTimer callback, AddressOf bridge, cleanup |

`ReDimHost.bas` exists because Excel can only call standard-module procedures from `Shape.OnAction`
and `SetTimer` callbacks. ROneCOne ADR 0007 declined a background completion pump for exactly this
reason under its one-file invariant and anticipated an external host loop. ReDim is that host.

## Component model

Retained mode. A component is a `ReDimUI` instance holding desired props (text, geometry, colors,
visible, enabled, value) plus the id of its worksheet shape(s). Rendering diffs desired props
against last-applied props and touches only changed shape members. A flush that repaints several
components, or opens or closes a drop list or calendar, runs with `ScreenUpdating` off, so Excel
paints the result once instead of part by part.

A control hidden before its first draw, as one on a tab not shown is, draws when it first
shows. The app's first render then draws what shows and nothing else: drawn hidden, the other
panels of the What's New demo's tabs were four fifths of its first render. On a sheet an earlier
session drew for the app, a hidden control draws as before, so a shape saved showing hides.

A control drawn from many parts, such as a list, table, picker, tab strip, or float field, hides
by hiding every shape it drew in one `ShapeRange` call, and shows the same shapes again in
another. Drawn part by part both ways, a later switch to the What's New demo's Lists tab took
76 ms; it takes 4. While hidden the control draws nothing but what it carries: a tab strip still
hides its panels, and a stack's member still tells its stack, which lays out again without it.
Shown again, it draws whatever changed meanwhile against the look it hid with. If a shape went
missing meanwhile, the others show one by one and every part draws again. A control hides part
by part as before while a list or calendar of its own is open, and on a sheet that held the
app's shapes when it mounted, since the app knows each shape it drew only on a sheet that held
none.

A control resolves its rectangle as it draws: from a cell, its own points, its stack's slot, or
the control it sits `Below` or `RightOf`. A resolved rectangle stands until the app hears of a
change (any control marked dirty, or a stack laying out again), so a chain of relative controls
resolves each link once, from the chain's far end, rather than each draw resolving every link
before it, nested one inside the next. A draw that moves a control queues the controls placed
against it, and the draw already running takes them in turn, so a move down a long chain nests
no draws either; both had run out of stack space near 200 links.

Shape naming: `rdm_<appId>_<componentId>` for a component's main shape, and
`rdm_<appId>_<componentId>__<part>` for the parts of composite widgets, such as a toggle's knob,
a list's rows, or a calendar's weeks. `Mount` is idempotent: an existing shape with a matching
name is adopted, so re-running setup code never duplicates shapes and app code can be re-entered
safely after a crash. A control starts knowing nothing of the parts an earlier session drew, so
the first render on such a sheet sweeps them: before anything draws, the parts only a focus, an
open list, or a tooltip has, and after, a list's rows numbered past what it now draws.

Every control is drawn from shapes; native form controls went in 0.5.0. Composite widgets keep
the look each part was last drawn with and rewrite only the parts whose look changed, and a part
whose place and font held rewrites only its text, fill, and ink. Creating a shape and writing its
look costs more than anything else a draw does, so parts are few and made cheaply. A list that
opens draws its first row in full and copies it (`Shape.Duplicate`) for the rest, since a copy
carries the size, border, margins, tab stops, font, and click. A calendar draws a week as one
line of text whose center tab stops sit over the day columns: six week lines and a mark shape
each for the date held, today, the key cursor, and the day under the pointer, where it once
drew a shape per day. Clicks on a week resolve to a day by the pointer's position. [api.md](api.md) describes the
widget set: buttons, labels, cards, progress bars, spinners, skeletons, toggles, tick boxes,
radio groups, steppers, sliders, selects, menu buttons, combos, transfer lists, check lists,
text fields, date pickers, images, tab strips, expanders, tables, badges, sparklines, toasts,
and the modal overlay.

## State

Each app owns a store mapping `String` keys to `Variant` values.

- `ui.SetState key, value` writes, marks bound components dirty, and flushes unless batched
- `ui.State(key)` reads
- Bindings: `.BindText key`, optional format applied through ROneCOne composite formatting,
  `.BindValue`, `.BindVisible`, `.BindEnabled`
- `ui.OnStateChanged key, "Module.Proc"` registers a zero-argument listener, run after a
  write changes the key and its bound controls redraw; a write that leaves the value as it
  was runs none

## Event dispatch

Every interactive shape's `OnAction` targets one dispatcher in `ReDimHost`. The dispatcher reads
`Application.Caller`, parses app and component ids, applies click guards (disabled, busy, debounce),
sets `ReDimUI.Sender` context, and runs the component's handler, the workbook procedure its
`OnClick` or `OnChange` names, through `Application.Run`; one it cannot run goes to the app's
`OnError` sink. An `OnClickAsync` body runs as a ROneCOne task through the pump. Tests can inject
clicks through the same path by calling the dispatcher with an explicit shape name.

## Async engine

The pump is a `SetTimer` callback (a 16 ms frame by default) plus a public `PumpOnce` for
deterministic tests. Each tick, inside a reentrancy guard with errors swallowed:

1. Advance animations: spinner rotation, toast slides, knob and tab-bar glides, a skeleton's
   pulse, caret blinks
2. Expire toasts
3. Watch what Shape macros cannot see: the pointer for hover looks, tooltips, hold-to-repeat,
   and drags, and presses or selection moves that close an open list or end a field's focus
4. Step registered async ops: call `AdvanceTask` (Friend, same-project) on the underlying ROneCOne
   task; on terminal state run done or fail handlers and restore bound controls
5. Run chunked jobs inside a per-tick millisecond budget using `GetTickCount64`

A frame visits only the controls that tick and reads the pointer at most once.

Task kinds and how they behave under the pump:

- Transport tasks (Delay, HTTP, ADO, process, file watch) genuinely overlap; each tick is one poll
- Delegate-run tasks execute their whole body in one step; long CPU work must use a Job
- Jobs call a step delegate repeatedly until it reports done, so cancel buttons and progress
  bars stay live during heavy loops

Sugar: `btn.OnClickAsync "Module.Proc"` disables the button, shows busy state, runs the work
through the pump, and restores on completion. Explicit form: `ui.Async(id)` builder with
`Disables`, `ShowsSpinner`, `OnDone`, `OnFail`, `OnCancel`, `TracksState`, `Start`. Cancellation
uses ROneCOne `CancellationTokenSource`.

Jobs run in two modes. Budget mode repeats the step inside a tick until the millisecond budget
elapses, for throughput work like imports. Paced mode runs at most one step per interval and may
be retuned live, for game loops and animations where cadence matters. The Snake demo exists to
prove the paced path.

Pump safety rails, in order of importance:

- The callback body is a single guarded call; no error ever escapes into Excel
- A consecutive-failure counter kills the timer after repeated faults
- The timer stops when no animations, ops, jobs, toasts, or watches remain
- `WorkbookBeforeClose` (Application events) and `ReDimUI.Shutdown` kill timers deterministically;
  a close leaves the apps mounted and arms nothing until the user acts in the workbook again,
  since the close can still be cancelled at the save prompt

## Theming

`ReDimUI.ThemeLight`, `ReDimUI.ThemeDark`, and `ReDimUI.ThemeHighContrast` presets, customized
with `WithPrimary` and `WithFont`. A theme carries the primary, surface, and muted colors with
their inks, the success, warning, danger, border, and canvas colors, and the font name and size;
`theme.ContrastReport` checks every pairing the controls draw against WCAG. `ui.SetTheme`
restyles every component through the normal diff path. A theme's revision follows its look,
shared by every theme whose tokens all match, so `SetTheme` with a fresh copy of the look
already drawn, as a rebuild does, restyles nothing. A new theme restyles a control's fills,
edges, and fonts; its words, text margins, visibility, click, and alternative text, which no
theme sets, stand as drawn, where rewriting them was half of a 100-control swap.

## Build and verification

- `tools/build_workbooks.py` injects sources into `.xlsm` files with pyOpenVBA and verifies
  byte-for-byte round trips, matching the ROneCOne pipeline. It removes the blank workbook's
  `Module1`, whose p-code held a comment its empty source did not, so a built workbook holds
  VBA source alone and Excel compiles it on open
- `tools/check.py` runs pyvbaanalysis over every `.bas` and `.cls`; any finding in ReDim's own
  sources fails the gate, and one inside the vendored ROneCOne is printed and not counted. CI
  installs the newest pyvbaanalysis on every run
- `tools/security_scan.py` runs olevba and mraptor over the runtime files and built workbooks
  and fails on any finding `tools/security_expected.json` does not list with its reason, and on
  any p-code in a workbook. The Security workflow runs it on every push, and the release
  security workflow runs it over each release's assets and attaches the report (see
  SECURITY.md)
- `tools/stamp_release.py` writes the release header into every source a release ships (both
  runtime files and each demo module): the version from `REDIM_VERSION`, that version's
  CHANGELOG date, the repository, and the MIT license text from `LICENSE`, ahead of
  `Option Explicit`. `tests/python/test_source_guards.py` fails until every source carries the
  current header, so a version bump cannot ship a stale one
- `tests/python/test_compile.py` compiles every shipped workbook with the real VBA compiler via
  pyvbaharness, because two grammar rules bit during development that static analysis does not
  model: statement-position calls with multiple parenthesized arguments, and case-insensitive
  locals shadowing same-named members
- `tests/python` drives live Excel through pyvbaharness: mount, render, dispatch, state,
  pump stepping via `PumpOnce`, async lifecycle, teardown
- `tests/python/test_casing.py` guards the host project's identifier casing: VBA keeps one
  spelling per name project-wide, so the runtime and demos may not declare a name in a casing
  that differs from the default type libraries or from ReDim's and ROneCOne's members, and a
  VBE export of ROneCOne, ReDim, every demo, and sample host code must return every token as
  written
- `tools/bench.py` times framework scenarios, the newer controls, and demo builds in live
  Excel, and compares runs saved with `--save`
- Demos are smoke-run live before release

Harness constraint, pinned by the spike suite: module globals do not survive across
`run_macro` round trips, so every live scenario completes inside one VBA call that returns a
transcript, and no SetTimer stays armed across harness call boundaries (the reset makes the
TIMERPROC address stale and the next WM_TIMER kills the process). Production Excel does not
mutate the project between interactions, so the constraint is test-only. The framework still
records the armed timer id in a workbook-scoped name and kills any orphan before arming a new
timer, which covers real state-loss events such as an unhandled error ending execution.

## Known constraints

- One Excel thread: pump ticks fire only while Excel idles or pumps messages; they pause during
  another macro, cell edit mode, or a modal dialog. This is documented, not hidden.
- `AdvanceTask` is Friend scope, which requires ReDim modules to live in the same VBA project as
  `ROneCOne.cls`. That is already the installation model.
