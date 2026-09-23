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
against last-applied props and touches only changed shape members inside a `ScreenUpdating` batch.

Shape naming: `rdm_<appId>_<componentId>` for a component's main shape, and
`rdm_<appId>_<componentId>__<part>` for the parts of composite widgets, such as a toggle's knob,
a list's rows, or a calendar's days. `Mount` is idempotent: an existing shape with a matching
name is adopted, so re-running setup code never duplicates shapes and app code can be re-entered
safely after a crash.

Every control is drawn from shapes; native form controls went in 0.5.0. Composite widgets keep
the look each part was last drawn with and rewrite only the parts whose look changed, and a part
whose place and font held rewrites only its text, fill, and ink. [api.md](api.md) describes the
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
- `ui.OnStateChanged key, "Module.Proc"` registers a handler (Action taking the key name)

## Event dispatch

Every interactive shape's `OnAction` targets one dispatcher in `ReDimHost`. The dispatcher reads
`Application.Caller`, parses app and component ids, applies click guards (disabled, busy, debounce),
sets `ReDimUI.Sender` context, and invokes the component's handler delegate (a ROneCOne Action,
typically wrapping a workbook procedure name). Tests can inject clicks through the same path by
calling the dispatcher with an explicit shape name.

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
- `WorkbookBeforeClose` (Application events) and `ReDimUI.Shutdown` kill timers deterministically

## Theming

`ReDimUI.ThemeLight`, `ReDimUI.ThemeDark`, and `ReDimUI.ThemeHighContrast` presets, customized
with `WithPrimary` and `WithFont`. A theme carries the primary, surface, and muted colors with
their inks, the success, warning, danger, border, and canvas colors, and the font name and size;
`theme.ContrastReport` checks every pairing the controls draw against WCAG. `ui.SetTheme`
restyles every component through the normal diff path.

## Build and verification

- `tools/build_workbooks.py` injects sources into `.xlsm` files with pyOpenVBA and verifies
  byte-for-byte round trips, matching the ROneCOne pipeline
- `tools/check.py` runs pyvbaanalysis over every `.bas` and `.cls`; any finding fails the gate
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
