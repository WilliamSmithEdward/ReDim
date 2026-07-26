# ReDim

A stateful UI framework for Excel worksheets. Build interactive, async-aware applications out of
worksheet shapes and form controls the way you would with a component framework on the web. No
UserForms, no ActiveX, no add-ins, no VBIDE access.

ReDim sits on top of [ROneCOne](https://github.com/WilliamSmithEdward/ROneCOne), which supplies
delegates, typed events, structured exceptions, and cooperative tasks. ReDim adds what ROneCOne's
one-file invariant cannot carry: the host loop (ROneCOne ADR 0007 anticipated exactly this), a
retained component model with diffed rendering, a state store with bindings, and an async engine
that keeps Excel responsive while work runs.

```vba
Dim app As ReDimUI

Set app = ReDimUI.Mount(Sheet1, "Dashboard")

app.Button("run").At("B2:C3").Text("Run Report").Primary.OnClickAsync "Demo.BuildReport"
app.ProgressBar("prg").At("B5:F5").BindValue "progress"
app.Label("status").At("B7:F7").BindText "statusMsg", "Status: {0}"

app.Render
```

Clicking Run disables the button, swaps its caption to busy text, runs `Demo.BuildReport` through
the pump, and restores the button when the work completes or fails. While a chunked job runs,
`app.SetState "progress", 40` moves the bar and `"statusMsg"` updates the label, all while the
user keeps working in the workbook.

## What it does

- **Components over shapes.** Button, Label, Card, ProgressBar, Spinner, Toggle, Checkbox,
  Dropdown, SelectBox, Slider, cell-backed TextInput, Toast, and a shapes-based modal overlay
  with Confirm. Drawn widgets carry a theme (light and dark presets). Where Excel fixes form
  control fonts, ReDim compensates: checkbox captions render as themed label parts that still
  toggle the box, and SelectBox is a fully drawn picker for when the native dropdown's small
  fixed font will not do.
- **Stateful.** Each app owns a key-value store. `BindText`, `BindValue`, `BindVisible`, and
  `BindEnabled` re-render only what changed. `OnStateChanged` registers workbook procedures as
  state listeners. `WritesTo` flows control values back into state. `Persist True` writes state
  through to hidden workbook names, so a closed and reopened workbook resumes exactly where the
  user left off; `SetStateDefault` keeps setup code from clobbering persisted choices.
- **Composable layout.** Anchor to ranges with `At`, to points with `AtRect`, or to other
  components with `Below` and `RightOf`. Render prunes shapes orphaned by renamed components,
  so iterating on app code never litters the sheet.
- **Windows.** Sheets as forms: `AsWindow` registers a window, `Navigate` shows one at a time
  (others go very-hidden), `NavigateBack` walks the stack, `OnShow` and `OnHide` fire like
  form lifecycle events, a button becomes a nav link with `NavigatesTo`, and `NavBar` renders
  window tabs with the active one highlighted. Background windows keep pumping.
  `ProtectSurface` turns an app sheet into a true application surface.
- **Genuinely async.** A SetTimer pump ticks whenever Excel idles and steps ROneCOne tasks
  without anyone blocking on `Await`. Fire-and-forget ops disable their controls, spin a spinner,
  and fire done, fail, or cancel handlers. HTTP, database, shell, and delay tasks overlap for
  real; CPU-bound VBA runs as chunked jobs inside a per-tick millisecond budget, or paced jobs
  for game loops and animations. Cancellation rides ROneCOne CancellationTokenSource.
- **Crash-railed.** The timer callback is a single guarded call that never lets an error escape,
  kills itself after repeated faults, stops when idle, records its timer id in a workbook name so
  orphans from VBA state loss are killed before re-arming, and tears down on workbook close.
- **Idempotent.** Re-running setup code adopts existing shapes instead of duplicating them, so an
  unhandled error that resets VBA state never leaves a broken sheet.

## Installation

1. Import `ROneCOne.cls` (from the ROneCOne release) into your macro-enabled workbook.
2. Import `src/ReDimUI.cls`.
3. Import `src/ReDimHost.bas`.

Three imports, no references, no registration. Windows x64 Microsoft 365 Excel, the same target
as ROneCOne.

## Demos

Built workbooks live in `demo/` after running `python tools/build_workbooks.py`:

| Workbook | Shows | Entry macro |
|---|---|---|
| `ReDim_Mission_Control.xlsm` | Three cancellable feeds with live progress, toasts, KPI cards, dark mode, confirm modal | `BuildMissionControl` |
| `ReDim_Widget_Gallery.xlsm` | Every component wired to a live state inspector | `BuildWidgetGallery` |
| `ReDim_Snake.xlsm` | A real game loop: paced pump job, arrow-key steering, score state, game-over modal | `BuildSnake` |
| `ReDim_Navigator.xlsm` | Sheets as forms: Navigate, back stack, lifecycle hooks, persisted settings | `BuildNavigator` |

## Documentation

- [Getting started and API](docs/api.md)
- [The async engine](docs/async.md)
- [Architecture and design contract](docs/design.md)

## Development

VBA sources are plain files; workbooks are built artifacts.

```bash
pip install -r requirements-dev.txt
python tools/check.py          # pyvbaanalysis static gate, zero findings required
python tools/build_workbooks.py
python -m pytest tests/python  # live Excel suite via pyvbaharness
```

The live suite covers mount and adoption, dispatch and guards, bindings, batching, theming, the
async op lifecycle, jobs, cancellation, every widget, a real armed-timer end-to-end run, a VBA
compile gate for every shipped workbook, and smoke runs of all three demos.

## License

MIT
