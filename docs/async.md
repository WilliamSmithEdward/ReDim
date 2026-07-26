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
With app.Async("refresh")
    .RunsTask ROneCOne.HttpClient.GetStringAsync("https://example.com/data")
    .Disables "btnRefresh"
    .ShowsSpinner "spn"
    .TracksState "refreshStatus"    ' running, done, failed, canceled
    .OnDone "Demo.ApplyData"
    .OnFail "Demo.ShowError"
End With
app.Async("refresh").Start
```

`RunsProc "Module.Proc"` wraps a workbook procedure instead of a task. `WithCancellation` gives
the op a ROneCOne token source: `app.CancelAsync "refresh"` cancels, `Token` exposes the token
for task factories such as `ROneCOne.Task.Delay(5000, op.Token)`. Disabled controls show busy
state (buttons swap to `BusyText`); everything restores on any terminal state. `AsyncError(opId)`
returns the failure message after a fault.

`btn.OnClickAsync "Module.Proc"` is the one-line form: a per-button op that disables the button,
runs the procedure, and restores it. Clicks while busy are ignored.

## Jobs

CPU-bound work stays responsive by running in slices:

```vba
app.Job("import").Steps("Demo.ImportChunk").BudgetMs(15) _
    .JobOnDone("Demo.ImportFinished").JobOnFail "Demo.ImportFailed"
app.Job("import").StartJob
```

The step procedure is a zero-argument `Function` returning `True` when finished. Budget mode
repeats the step inside each tick until the budget elapses. Paced mode runs at most one step per
interval, the right shape for game loops and animations, and the pace can change while running:

```vba
app.Job("loop").Steps("Game.Frame").PacedMs 150
```

`CancelJob` requests a stop; the job's cancel handler runs on the next tick. Step errors finish
the job through its fail handler and the app error sink instead of surfacing a dialog.

## Rails

- The timer callback body is one guarded call. Errors never escape into Excel; ten consecutive
  faults kill the timer.
- Renders that happen inside a tick never toggle `Application.ScreenUpdating`: flipping it at
  pump frequency redraws scroll bars and blips the cursor. Batched flushes outside ticks still
  use it.
- While the timer is armed, `Application.Cursor` is pinned to the arrow and restored to default
  on stop. Excel otherwise flips to the busy cursor on every VBA execution, which at pump
  frequency reads as a strobe. The steady arrow over the grid doubles as the "work is running"
  signal.
- While the timer is armed, the multimedia timer resolution is raised to 1 ms and restored on
  stop, so frames land on schedule instead of quantizing into the default ~15.6 ms buckets.
- While anything animates, budget jobs still using the default budget yield to the frame rate
  (8 ms per work pass); an explicit `BudgetMs` is always honored as given.
- The pump stops itself when no ops, jobs, toasts, or visible spinners remain, and re-arms on
  the next `Start`.
- The armed timer id is stored in a workbook-scoped name. After VBA state loss, the next arm
  kills the orphan first.
- `WorkbookBeforeClose` and `ReDimUI.Shutdown` stop the pump deterministically, so no TIMERPROC
  outlives its project.
- `ReDimUI.AutoPump False` plus `ReDimUI.PumpOnce` gives tests a fully deterministic clock.

## Honest limits

One thread. Ticks fire only while Excel pumps messages: they pause during another macro, while a
cell is being edited, and behind native modal dialogs. Delegate bodies block their tick for their
full duration. These are Excel's rules; ReDim's job is to make the cooperative model feel
effortless, not to pretend the host is multithreaded.

Pauses are visual, not corrupting: animations are time-based, paces are wall-clock, and task
state advances on the next tick, so everything catches up correctly when the interaction ends.
The one preventable stall is edit mode on the app sheet itself, and the opt-in `ProtectSurface`
closes it by treating that sheet as an application surface. ReDim deliberately does not reach
for `Application.Interactive` or dialog suppression: locking a user out of their own Excel to
keep a spinner smooth inverts who the session belongs to, and a failure while input is disabled
leaves Excel unusable.
