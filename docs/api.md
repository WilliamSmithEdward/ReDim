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
| `ReDimUI.AutoPump enabled` | Turn the wall-clock timer off for deterministic runs. |
| `ReDimUI.Shutdown` | Kill the pump and forget every app. Shapes stay. |
| `ReDimUI.ThemeLight`, `ThemeDark` | Theme presets; customize with `WithPrimary`, `WithFont`. |

App ids use letters and digits only. Component ids may add single underscores.

## App

Component factories, get-or-create by id: `Button`, `Label`, `Card`, `Spinner`, `ProgressBar`,
`Toggle`, `Checkbox`, `Dropdown`, `SelectBox`, `Slider`, `TextInput`. Also:

| Member | Purpose |
|---|---|
| `SetState key, value` / `State(key)` / `StateOrDefault(key, fallback)` / `HasState(key)` | The store. |
| `SetStateDefault key, value` | Sets only when the key has no value; the right form for initial values. |
| `Persist enabled` | Writes state through to hidden workbook names and hydrates persisted values, so a reopened workbook resumes where the user left off. Scalars only; call before Render. |
| `ClearPersisted` | Deletes this app's persisted names. |
| `HotKey keyCode, "Module.Proc"` / `ClearHotKeys` | Application.OnKey with cleanup on Unmount and Shutdown. |
| `OnStateChanged key, "Module.Proc"` | Zero-argument listener runs after the key changes. |
| `BeginUpdate` / `EndUpdate` | Batch several changes into one flush. |
| `Render` | Mark everything dirty and paint. Call once after building the UI. |
| `SetTheme theme` | Restyle every component. |
| `PrepareCanvas` | Paint the sheet background and hide gridlines. |
| `Toast message, ttlMs` | Transient card on a rail beside the content, expiring on the pump or on click. Slots reuse lowest-first; modal chrome never shifts the rail. |
| `ToastTray anchor` | Pins the tray's top-left to a range instead of the default rail. |
| `Confirm title, message, okProc, cancelProc, okText, cancelText` | Shapes-based modal. |
| `CloseModal` | Hide the modal set. |
| `Async(opId)` / `CancelAsync opId` / `AsyncError(opId)` | Async ops (see [async.md](async.md)). |
| `Job(jobId)` / `CancelJob jobId` | Chunked or paced background work. |
| `OnError "Module.Proc"` | One-argument sink for swallowed handler failures. |
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
- Values: `Value(number)` (progress, slider, dropdown index), `Checked(flag)`,
  `Items("A", "B", ...)`, `SliderRange(min, max, step)`.
- Bindings: `BindText(key, template)` where `{0}` is the value, `BindValue(key)`,
  `BindVisible(key, invert)`, `BindEnabled(key, invert)`, `WritesTo(key)`. The invert flag
  serves the disable-while-busy pattern: `BindEnabled "anyRunning", True`.
- Behavior: `OnClick "Module.Proc"`, `OnClickAsync "Module.Proc"`, `OnChange "Module.Proc"`.
- Reads: `CurrentValue`, `CurrentText`, `IsChecked`, `IsEnabled`, `IsBusy`, `InputValue`
  (TextInput; assigning it writes the cell without firing change events).
- `Remove` deletes the component and its shapes.

Handlers are zero-argument public procedures referenced as `"Module.Proc"`. Inside a handler,
`ReDimUI.Sender` carries who fired and `ReDimUI.SenderPart` names the clicked sub-shape when a
composite widget fired (toggle knob, checkbox caption, select option).

## Text sizing on controls

Excel fixes the fonts of form controls: checkbox captions and dropdown lists always render at
the small system size and expose no Font setting. ReDim compensates where it can:

- `Checkbox` keeps its native caption empty and draws the caption as a themed textbox part.
  Clicking the caption toggles the box, so behavior matches a native caption.
- `SelectBox` is a fully drawn picker (face, caret, option list) that follows the theme font.
  Prefer it when type consistency matters; keep `Dropdown` when the native Excel look is wanted.
- `Slider` has no text. Pair it with a bound `Label`, as the gallery does.

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
