# ReDim design

ReDim is a stateful UI framework for Excel worksheets. It renders retained components as worksheet
shapes and form controls, binds them to a per-app state store, and drives async behavior through a
timer pump layered over ROneCOne's cooperative task scheduler. No UserForms, no ActiveX, no VBIDE
access.

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
against last-applied props and touches only changed shape members inside a `ScreenUpdating` batch.

Shape naming: `rdm_<appId>_<componentId>` (plus suffixes such as `_fill` or `_knob` for composite
widgets). `Mount` is idempotent: an existing shape with a matching name is adopted, so re-running
setup code never duplicates shapes and app code can be re-entered safely after a crash.

Widget set, v1:

- Button: rounded rectangle, styled states (normal, hover-free, disabled, busy)
- Label: borderless text box
- Card: rectangle panel with optional title, used for KPI tiles and modal bodies
- ProgressBar: track rectangle plus fill rectangle, value 0 to 100
- Spinner: arc shape rotated by the pump while visible
- Toggle: drawn pill plus knob, flips a bound state key
- Checkbox, Dropdown, Slider: native form controls wrapped and synced to state keys
- TextInput: cell-backed input with a frame shape, change captured via Application events
- Toast: transient card with TTL, dismissed by the pump
- Overlay and Modal: dimming rectangle that swallows clicks plus a centered card with buttons

## State

Each app owns a store mapping `String` keys to `Variant` values.

- `app.SetState key, value` writes, marks bound components dirty, and flushes unless batched
- `app.State(key)` reads
- Bindings: `.BindText key`, optional format applied through ROneCOne composite formatting,
  `.BindValue`, `.BindVisible`, `.BindEnabled`
- `app.OnStateChanged key, "Module.Proc"` registers a handler (Action taking the key name)

## Event dispatch

Every interactive shape's `OnAction` targets one dispatcher in `ReDimHost`. The dispatcher reads
`Application.Caller`, parses app and component ids, applies click guards (disabled, busy, debounce),
sets `ReDimUI.Sender` context, and invokes the component's handler delegate (a ROneCOne Action,
typically wrapping a workbook procedure name). Tests can inject clicks through the same path by
calling the dispatcher with an explicit shape name.

## Async engine

The pump is a `SetTimer` callback (default 50 ms) plus a public `PumpOnce` for deterministic tests.
Each tick, inside a reentrancy guard with errors swallowed:

1. Advance animations: spinner rotation, indeterminate progress sweep
2. Expire toasts
3. Step registered async ops: call `AdvanceTask` (Friend, same-project) on the underlying ROneCOne
   task; on terminal state run done or fail handlers and restore bound controls
4. Run chunked jobs inside a per-tick millisecond budget using `GetTickCount64`

Task kinds and how they behave under the pump:

- Transport tasks (Delay, HTTP, ADO, process, file watch) genuinely overlap; each tick is one poll
- Delegate-run tasks execute their whole body in one step; long CPU work must use a Job
- Jobs call a step delegate repeatedly until it reports done, so cancel buttons and progress
  bars stay live during heavy loops

Sugar: `btn.OnClickAsync "Module.Proc"` disables the button, shows busy state, runs the work
through the pump, and restores on completion. Explicit form: `app.Async(id)` builder with
`Disables`, `ShowsSpinner`, `OnDone`, `OnFail`, `OnCancel`, `TracksState`, `Start`. Cancellation
uses ROneCOne `CancellationTokenSource`.

Jobs run in two modes. Budget mode repeats the step inside a tick until the millisecond budget
elapses, for throughput work like imports. Paced mode runs at most one step per interval and may
be retuned live, for game loops and animations where cadence matters. The Snake demo exists to
prove the paced path.

Pump safety rails, in order of importance:

- The callback body is a single guarded call; no error ever escapes into Excel
- A consecutive-failure counter kills the timer after repeated faults
- The timer stops when no animations, ops, jobs, or toasts remain
- `WorkbookBeforeClose` (Application events) and `ReDimUI.Shutdown` kill timers deterministically

## Theming

`ReDimUI.ThemeLight` and `ReDimUI.ThemeDark` presets plus a custom builder: primary, success,
danger, surface, border, text, muted colors, font name and size, corner radius. `app.SetTheme`
restyles every component through the normal diff path.

## Build and verification

- `tools/build_workbooks.py` injects sources into `.xlsm` files with pyOpenVBA and verifies
  byte-for-byte round trips, matching the ROneCOne pipeline
- `tools/check.py` runs pyvbaanalysis over every `.bas` and `.cls`; any finding fails the gate
- `tests/python/test_compile.py` compiles every shipped workbook with the real VBA compiler via
  pyvbaharness, because two grammar rules bit during development that static analysis does not
  model: statement-position calls with multiple parenthesized arguments, and case-insensitive
  locals shadowing same-named members
- `tests/python` drives live Excel through pyvbaharness: mount, render, dispatch, state,
  pump stepping via `PumpOnce`, async lifecycle, teardown
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
- Form controls carry Excel's native look; drawn widgets (Button, Toggle, ProgressBar) carry the
  theme.
