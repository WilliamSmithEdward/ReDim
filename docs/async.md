# The ReDim async engine

## Why this works

ROneCOne tasks are cooperative: whoever waits on a task advances it. `Await` keeps Excel alive
by pumping messages, but somebody still has to block. ROneCOne ADR 0007 declined the alternative,
a background completion pump, because the pump's timer callback can only live in a standard
module and ROneCOne ships as a single class by hard invariant. The ADR explicitly left the door
open for an external host loop.

ReDim is that host. `ReDimHost.bas` arms a `SetTimer` callback at a 16 ms animation frame rate.
Excel dispatches WM_TIMER whenever it pumps messages, which includes ordinary idle time between
user actions. Animations advance every frame with measured elapsed time; ops, budget jobs, and
toast expiry run on a 50 ms work cadence inside the same loop, so the higher frame rate buys
smoothness without multiplying task polling or job duty. Each work pass calls `AdvanceTask`
(Friend scope, same project) on every in-flight task, non-blockingly. The result is
fire-and-forget async in VBA: start work, return immediately, and let completion handlers fire
while the user keeps editing cells. `PumpOnce` advances one nominal 50 ms frame with the work
pass included, which is what keeps tests deterministic.

Transport tasks (Delay, HTTP, ADO, shell, file watch) genuinely overlap; each tick is one poll.
A delegate task runs its whole body inside one tick, so long CPU-bound VBA belongs in a job.

## Async ops

```vba
With ui.Async("refresh")
    .RunsTask ROneCOne.HttpClient.GetStringAsync("https://example.com/data")
    .Disables "btnRefresh"
    .ShowsSpinner "spn"
    .TracksState "refreshStatus"    ' running, done, failed, canceled
    .OnDone "Demo.ApplyData"
    .OnFail "Demo.ShowError"
End With
ui.Async("refresh").Start
```

`RunsProc "Module.Proc"` wraps a workbook procedure instead of a task. `WithCancellation` gives
the op a ROneCOne token source: `ui.CancelAsync "refresh"` cancels a running op (an op at rest
has nothing to cancel), `Token` exposes the token for task factories such as
`ROneCOne.Task.Delay(5000, op.Token)`. A cancelled op finishes with a fresh token, so its next
run, and a `Token` read for it, start uncancelled. Disabled controls show busy
state (buttons swap to `BusyText`); everything restores on any terminal state, and a spinner
or control that several running ops share stays busy until the last of them finishes. The
op's body runs inside the pump with the op as `ReDimUI.Sender`. `OnCancel` names
the handler a cancel runs, as `OnDone` and `OnFail` name theirs. `AsyncError(opId)` returns the
failure message after a fault. `ui.Async("refresh").IsRunning` says whether the op is running
now, so a second click can wait for the first run; `ui.Job(id).IsRunning` does the same for a
job. Cancelling an op or job that was never created does nothing, as it does for one at rest.

An outcome handler runs with the op as `ReDimUI.Sender`, so it reads the result without a
module variable holding the task. Like every handler it traps its own errors: Excel does not
hand an error raised inside a procedure `Application.Run` started back to ReDim, so an
untrapped one stops the pump's tick with VBA's dialog. A handler ReDim cannot run goes to the
app's `OnError` sink.

```vba
Public Sub ApplyData()
    Dim body As String
    body = ReDimUI.Sender.Task.Result
    ' ...
End Sub
```

`Task` is the ROneCOne task the op runs, or Nothing before it starts. A job's `JobOnDone`,
`JobOnFail`, and `JobOnCancel` handlers run with the job as the sender the same way.

`btn.OnClickAsync "Module.Proc"` is the one-line form: a per-button op that disables the button,
runs the procedure, and restores it. Clicks while busy are ignored. The procedure runs with the
button as `ReDimUI.Sender`, as a click handler does, so buttons told apart by their `Tag` can
share it.

## Jobs

CPU-bound work stays responsive by running in slices:

```vba
ui.Job("import").Steps("Demo.ImportChunk").BudgetMs(15) _
    .JobOnDone("Demo.ImportFinished").JobOnFail "Demo.ImportFailed"
ui.Job("import").StartJob
```

The step procedure is a zero-argument `Function` returning `True` when finished. It runs with
the job as `ReDimUI.Sender`, so a step shared by several jobs tells them apart by the job's
`Tag`, as the Mission Control demo's three feeds do. Budget mode
repeats the step inside each tick until the budget elapses. Paced mode runs at most one step per
interval, the right shape for game loops and animations, and the pace can change while running:

```vba
ui.Job("loop").Steps("Game.Frame").PacedMs 150
```

`CancelJob` requests a stop; the job's `JobOnCancel` handler runs on the next tick. A step
ReDim cannot run finishes the job through its fail handler and the app's error sink. A step
traps its own errors, as a handler does.

## Rails

- The timer callback body is one guarded call. Errors never escape into Excel; ten consecutive
  faults kill the timer.
- The timer callback is VBA code, and a WM_TIMER that reaches it while VBA is in break mode, or
  after an edit to `ReDimHost`, takes Excel down. So the timer is held, set to wait 24 days with
  its id kept, while code ReDim runs could stop in the debugger: each tick, each click, key,
  shortcut, and Excel event ReDim handles, and each handler, step, listener, check, and rebuild
  it calls. The last hold to let go sets the next frame. A breakpoint, `Stop`, or Debug in VBA's
  error dialog anywhere in that code finds no WM_TIMER due.
- The pump pauses while Excel's active window is not a workbook's own: the VBA editor or one of
  its windows, or a VBA dialog such as a run-time error or a `MsgBox`, where an edit or a break
  is a keystroke away. Each tick checks, and so does each arm and each hold letting go, which
  covers stepping out of a handler in the debugger. A paused pump comes back at the next click,
  key, selection, or sheet change in the workbook. A dialog Excel owns, such as Format Cells,
  does not pause it, and neither does a hidden Excel with no window active.
- Renders that happen inside a tick never toggle `Application.ScreenUpdating`: flipping it at
  pump frequency redraws scroll bars and blips the cursor. Flushes outside ticks still use it
  when they repaint several controls, or open or close a drop list or calendar.
- Cursor pinning is opt-in (`ReDimUI.PinPumpCursor True`). Pinning suppresses Excel's
  busy-cursor flip during each tick, but it also overrides context cursors, hiding the hover
  hand on interactive shapes while work runs. Since ticks became lean the strobe is
  negligible, so hover affordance wins by default.
- While the timer is armed, the multimedia timer resolution is raised to 1 ms and restored on
  stop, so frames land on schedule instead of quantizing into the default ~15.6 ms buckets.
- While anything animates, budget jobs still using the default budget yield to the frame rate
  (8 ms per work pass); an explicit `BudgetMs` is always honored as given.
- The pump stops itself once nothing needs it: no op, job, toast, or animation, and no watch
  for a focused control, the pointer, a press, or a drag. It arms again when one starts.
- The armed timer id is stored in a workbook-scoped name. After VBA state loss, the next arm
  kills the orphan first, and a tick from a timer the current state did not arm kills that
  timer and gives back the timer resolution its arming raised.
- `WorkbookBeforeClose` and `ReDimUI.Shutdown` stop the pump deterministically, so no TIMERPROC
  outlives its project. A close keeps the apps mounted in case it is cancelled at the save
  prompt; nothing arms the pump again until the user clicks or selects in the workbook.
- An op whose tick faults does not hold up the ops after it, and a `TracksState` listener
  ReDim cannot run does not cost the op its outcome handler; the failure goes to the app's
  error sink.
- `ReDimUI.AutoPump False` plus `ReDimUI.PumpOnce` gives tests a fully deterministic clock.

## Honest limits

One thread. Ticks fire only while Excel pumps messages: they pause during another macro, while a
cell is being edited, and behind native modal dialogs. Delegate bodies block their tick for their
full duration. These are Excel's rules; ReDim's job is to make the cooperative model feel
effortless, not to pretend the host is multithreaded.

While a click or key handler runs, the pump waits, even when the handler waits inside itself
with `DoEvents`, a blocking `Await`, or a `MsgBox`: a handler that loops until an op finishes
never sees it finish. The op's outcome handlers are the way to wait on it. The debugger rails
cover the code ReDim runs. Code it does not start, run from Excel while the pump runs, such as a
macro from a Forms button or the Macros dialog or a sheet event of the workbook's own, can still
stop at a breakpoint with a WM_TIMER due, and Excel closes. Nothing of ReDim's runs between that
macro starting and its stop. Run such code from the VBA editor, where the pump waits.

Pauses are visual, not corrupting: animations are time-based, paces are wall-clock, and task
state advances on the next tick, so everything catches up correctly when the interaction ends.
The one preventable stall is edit mode on the app sheet itself, and the opt-in `ProtectSurface`
closes it by treating that sheet as an application surface. ReDim deliberately does not reach
for `Application.Interactive` or dialog suppression: locking a user out of their own Excel to
keep a spinner smooth inverts who the session belongs to, and a failure while input is disabled
leaves Excel unusable.
