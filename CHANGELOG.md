# Changelog

## 1.0.3 - 2026-09-25

- `WritesTo` goes both ways: a control follows the key it writes.
  `SetState` on the key moves the control, firing no `OnChange`, so New
  game in the Snake demo turns its Pause switch off with the state, where
  the switch stayed on. On the first render a key with no value takes
  the control's value, as `SetStateDefault` would, so a build no longer
  seeds each key by hand; a key that has one shows on the control, so a
  rebuild keeps the value in play. A select, radio group, or tab strip
  maps the item's value or text back to the item, and a field that has
  the keys keeps what is being typed. Check lists, transfer lists, and
  cell-backed fields do not follow. `BindValue` on a select, radio
  group, or tab strip now takes an item's text as well as a position.
  The demos drop their seeding lines and hand-written loads, and the
  Expense Tracker's budget stays a number instead of turning into text.
  Every write of the key moves the control, even to the value it
  already held, so a switch changed with `Checked` goes back when code
  writes the key; a pick among items sharing a value stays on the item
  picked; `Empty` and `Null` clear a pick; a date comes from a serial
  number or a numeric string; state filling a `Required` pick clears
  its message; and a label bound to a seeded key shows the seed, even
  one drawn before the control that seeds it.
- An `OnStateChanged` listener runs after the controls bound to the key
  redraw, so one reading `SelectedText` or `IsChecked` sees the new
  value rather than the old one.
- Bindings match state keys in any case, as the state store does:
  `WritesTo "Paused"` followed by `SetState "paused"` updated the value
  but never redrew the control.
- A control whose draw raises no longer stops the app from drawing. Its
  id stayed in the dirty list, so every later `SetState` raised again
  and the controls after it never drew until the app was rebuilt.
- Typing can no longer stay dead after an error. A handler that raised
  while focus moved between fields left a flag up that kept every later
  blur from releasing the keys, and a reset of the VBA project under a
  focused field kept them captured with nothing focused; a key arriving
  with nothing focused now gives them back.
- A focused control takes the typing keys only while its sheet is in
  front. A `Confirm` raised from an op on a sheet behind another took
  the keys, so typing on the sheet in view did nothing and Enter
  answered a dialog no one could see.
- A `Confirm` opened from a dialog's OK or Cancel handler stays open:
  the button now closes its dialog before running the handler, where it
  closed the new one afterwards. Focus returns past both dialogs to the
  control that held it before the first.
- A focused field that hides or disables commits what was typed, running
  `OnChange`, `WritesTo`, and its checks; the text was dropped.
- Cancelling a workbook close at Excel's save prompt no longer leaves a
  dead UI. The close forgot every app, so clicks did nothing until the
  UI was rebuilt; it now stops the pump and gives the keys back, keeps
  the apps, and resumes at the next click or selection in the workbook.
- A cancelled async op runs again. Its token stayed cancelled, so every
  later `Start` was cancelled at once; `CancelAsync` on an op at rest
  now does nothing, and a finished cancel leaves a fresh token.
- A `TracksState` listener that raises no longer costs an op its
  `OnDone`, `OnFail`, or `OnCancel`, and an op whose tick faults no
  longer holds up the ops after it.
- `Unmount` stops the app's running ops and jobs, where they kept their
  controls busy and a job step that unmounted its own app kept looping
  against it.
- Removing a control that others sit `Below` or `RightOf` leaves them
  where they are; every later draw of theirs raised for want of the
  anchor. `At` after `Below` or `RightOf` now replaces the relative
  placement instead of drawing beside the old anchor.
- `InputValue` on a cell-backed field turns Excel's events back on when
  the cell write fails, where they stayed off for the session.
- Word moves and deletes treat CJK and Hangul text as letters. Characters
  from U+8000 up read as punctuation, so Ctrl+Backspace stopped at every
  character.
- A component id may not start or end with an underscore, which made its
  part names split at the wrong place: clicks reached the wrong control
  and `Render` swept its parts as leftovers.
- A `Toggle` with a caption in a `Stack` running across takes the
  caption's room, where the next member drew over the caption.
- `Async(id).IsRunning` and `Job(id).IsRunning` read whether an op or
  job runs now; the demos used undocumented Friend members, which a
  project referencing ReDim cannot call. `CancelAsync` and `CancelJob`
  on an id never created do nothing, so a reset that cancels every job
  needs no error trap.
- `Tag(value)` attaches a value to a control, op, or job, read back as
  `ReDimUI.Sender.TagValue`; a job's step now runs with the job as the
  sender. The Mission Control demo's twelve one-line feed handlers
  become four.
- A handler failure names the procedure and what ran it ("ReDim handler
  error 13 in Orders.Save for component 'save': ..."), and without an
  `OnError` sink it also goes to the Immediate window, since the status
  bar loses it at the next message. An `OnChange` that fails from a
  field's commit, a slider's release, or a transfer list's move now
  reaches the sink too, with the control as `ReDimUI.Sender`; those
  paths dropped the error without a trace.
- A `RadioGroup` takes Page Up, Page Down, and letters, as a `SelectBox`
  does, and a `SlideBar` takes typed digits, as a `Stepper` does.
- A `CheckList` or `TransferList` seeds its `WritesTo` key with its
  checks or picks on the first render, as the other controls do; the
  Widget Gallery drops its two hand-written seeds.
- `ItemPosition(text)` finds an item by its text and `ChosenPosition(text)`
  a transfer list's chosen item; the ReDex demo's catch check loses its
  two search loops.
- A control that cannot take its key's value yet takes it later: a
  focused field once the user leaves it, a select once its item loads;
  the first try was the only one. An `OnStateChanged` listener that
  rewrites a key a control just wrote now moves the control back, a pick
  among items sharing a value keeps its place through `Render` and theme
  changes, and `WritesTo` and `BindValue` keys match in any case.
- A draw that raises during `SetState` no longer skips the key's
  listeners; the error still comes out after them.
- Focus moving to a sheet behind gives back the shortcuts that share the
  typing keys, such as Ctrl+Z, which went to Excel instead. A failure
  in the field left behind goes to the app's error sink, and the control
  gaining focus still draws as focused.
- `ReDimUI.SenderId` inside a job's step names the job; it raised there,
  failing the job. A paced job whose step unmounts its app stops without
  running `JobOnDone` against the app that is gone.
- A `SlideBar` fires no `OnChange` for a click or a typed digit that
  snaps back to the value it held.
- A table's rows and headers come back whole after a shrink or a
  filter's footer dropped them; they stayed missing, or drew at the
  table's corner, until the data changed.
- `BindText`, `BindVisible`, `BindEnabled`, and `BindSource` read Null
  as nothing or False, where they raised error 94 and stopped the draw,
  and text as True when it holds anything.
- An `Image` whose source changes to a missing file, or to none, shows
  the placeholder, not the last picture; a picture saved with the
  workbook still outlives its own source. The placeholder takes a new
  theme's colors, and the file check no longer resets a `Dir` loop the
  caller is walking.
- `ExportTo` writes text as text: "00123" keeps its zeros and "=A1"
  stays words, where Excel turned them into a number and a live formula.
- A live `Sparkline` reshapes its line in place, so it stays under an
  open drop list or a dialog; it drew over them, and its clicks reached
  it through the dialog.
- The command palette stays shut while a `Confirm` waits, as the app's
  shortcuts do; it opened over the dialog and ran the buttons behind it.
- `Navigate` takes an app id in any case. A different case hid the very
  sheet it showed and fired OnHide and OnShow on the same window.
- A `Badge`'s tooltip shows over its words, and a `Toggle`'s caption
  counts as part of the switch for the pointer.
- A toast no longer takes an app's own `toast_1`, and ids starting with
  `mdl_` are ReDim's own dialog parts: making one raises a clear error.
- A stretched `Stack` member takes its own width back when the stretch
  ends or it leaves the stack, a removed stack's members show again and
  go free, `ColumnFormat` filters anew, and a control placed `Below` a
  component spelled in another case follows its growth.
- `Items` and `ItemsFrom` keep a check list's checks on the items still
  there, by text. `ItemsFrom` cleared them, and `Items` kept them by
  position, on whatever item came to sit there.
- A check list's boxes take the pressed look a tick box has.
- Typing a filter into a scrolled check list brings its first match
  into view; the window stayed where it was, over rows the filter hid.
  A check list's filter also reads its rows from the cache a transfer
  panel uses, where each key press scanned every item up to five times.
- Dead code goes: the unused `RdxTickCount` in `ReDimHost` and its
  counter, an unused toast reader, parameters two procedures never read
  (`CurrentText` looked up its app for nothing), and five spike-test
  helpers no test calls.
- A one-line face cut short ends in an ellipsis: an unfocused
  `TextInput` or `ComboBox`, a `SelectBox`, and a `DatePicker`. They
  stopped mid-word on one line, so "A very long selection" read as
  "A very long" and a long date format as "Wednesday". The face is
  measured as drawn, and the value keeps its whole text.
- A theme's other colors have builders: `WithSurface`, `WithMuted`,
  `WithStatus`, `WithBorder`, and `WithCanvas` join `WithPrimary` and
  `WithFont`, so a custom theme no longer needs the undocumented Friend
  `ConfigureTheme`. The ReDex demo's theme is built with them.
- A focused `Stepper` takes typed digits: typed within a second of
  each other they set its value, clamped to the range, as an up-down
  control's edit box does. Reaching 250 took 25 Page Ups.
- A text field's builders raise on a control they do not apply to:
  `Placeholder`, `OnInput`, `DebounceMs`, `Numeric`, `MaxLength`,
  `Validates`, `Clearable`, `Masked`, and `ErrorText` name the kinds
  they take, where they did nothing on, say, a Toggle, as `MultiLine`
  and `RestrictToItems` already raised. `ErrorText` shows under a
  `SelectBox`, `DatePicker`, or `RadioGroup` whether or not it is
  `Required`; it showed only on a required one.
- `OnStateChanged` takes an array of keys, `Array("volume", "mode")`,
  to put one listener on each, and adds a listener to a key only once.
  A build that ran twice in a session registered every listener again,
  so each change fired it twice. The Widget Gallery and ReDex demos
  wire their listeners in one call each.
- An async op's `OnDone`, `OnFail`, and `OnCancel` handlers, and a
  job's, run with the op or job as `ReDimUI.Sender`, and an error
  inside one reaches the app's `OnError` sink; they ran under `On Error
  Resume Next`, so a failure there vanished. The new `Task` reader
  gives the op's ROneCOne task, so `ReDimUI.Sender.Task.Result` reads
  the result without keeping the task in a module variable.
- Keys a Windows text box and tab strip take now work too. Caps Lock
  turns a typed letter's case over; OnKey binds keys, not characters,
  so letters typed with Caps Lock on came out lowercase. Shift+Backspace
  deletes where it went to Excel, and Ctrl+Ins, Shift+Del, and
  Shift+Ins copy, cut, and paste. Ctrl+Page Up and Ctrl+Page Down step
  a focused tab strip and a calendar's year, and no longer switch the
  sheet while a control has the keys. A `Filterable` list that matches
  nothing says "No matches", not "No items".
- `Filterable` lets typing filter a `SelectBox`'s or `MenuButton`'s
  open list, as a combo's list filters: the face shows the letters,
  the matches bold them, and the arrows, Enter, and clicks work among
  the matches. Backspace takes a letter off and Esc clears them before
  it closes the list. A letter on the closed list opens it, and a click
  that opens it gives it the keys. Off by default, where letters still
  jump to an item.
- A long `CheckList` or `RadioGroup` scrolls. Every row used to share
  the control's height however many items there were, so 30 items in
  150 points made rows under 5 points: boxes overlapped, captions could
  not be read, and click targets fell far below 24 pixels. Rows now
  never go under 18 points; a list whose rows would shows a window of
  rows with paging arrows at its right edge, which page it, repeat
  while held, and read muted at the ends. The keys move the window with
  the cursor or selection, and Page Up and Page Down page a windowed
  check list instead of jumping to its ends. Lists whose rows fit draw
  as before.
- A `Toggle`'s `Text` draws as a caption right of the switch that flips
  it when clicked and names it for screen readers ("Dark mode, Switch,
  on"). `Text` did nothing on a Toggle, so every demo placed a separate
  `Label` beside each switch; the six demos now caption their switches
  instead, and a click on the words flips the switch.
- `Required` on a `SelectBox`, `DatePicker`, or `RadioGroup` is checked:
  `ValidateAll` fails while nothing is picked, shows the message under
  the control, and puts focus on it, and a pick takes the message down.
  A `Stack` leaves room under the control for the message.
  It used to draw only the caption's asterisk, and `ValidateAll` passed
  a form with nothing picked.
- `ItemEnabled` works on a `ComboBox` and a `RadioGroup`, where it used
  to raise: a disabled item reads muted, clicks and keys pass over it,
  and a combo never suggests it.
- `SelectedText` and `SelectedValue` read the selection of a
  `SelectBox`, `ComboBox`, `RadioGroup`, or `Tabs` the same way: the
  item's text and value, or empty while nothing is selected.
  `CurrentText` answers a SelectBox's placeholder when nothing is
  picked, so code that saved it could store the placeholder as data.
- Buttons with nowhere to go read muted. A `Stepper`'s minus at its
  minimum and plus at its maximum, and a `TransferList` panel's up
  arrow on its first page and down arrow on its last, draw their glyph
  in muted ink, as a drop list's pager at the end of its list does. A
  click on them changed nothing before and still changes nothing, but
  looked live. A `TransferList`'s move buttons read muted the same way:
  `>` and `<` while their panel has no selection, `>>` and `<<` while
  it shows no rows.
- The ReDex demo no longer catches a Pokemon twice once it is in the
  party. Its duplicate check read only the box side of the team list.
- A control's `Caption` leads its alternative text, so a screen reader
  names the field: "Email, Edit field", where it read "Edit field".
- `Placeholder` works on a `SelectBox` and a `DatePicker`, as the same
  words given with `Text`. It used to do nothing there.
- A long selection, date, or typed text no longer runs under the drop
  caret of a `SelectBox`, `DatePicker`, or `ComboBox`. The face text
  kept a 6-point right margin, while the caret sits 12 to 14 points in
  from the edge; a combo made room for it only while its clear button
  showed. The margin now clears the caret, and a combo's overflow
  window counts it.
- Ctrl+A in a focused `CheckList` checks every row that shows and
  fires `OnChange` once, and in a `TransferList` selects every row of
  the cursor's panel. ReDim captured the key while such a list had
  focus and did nothing with it. A second Ctrl+A clears nothing, as in
  a native list.
- A `Table`'s footer pager and a `DatePicker` calendar's month arrows
  repeat while held, as the stepper, drop list, and transfer arrows
  already did, so a long table or a date months away takes one press
  instead of a click per page or month.
- A `RadioGroup` or `CheckList` emptied by `ClearItems`, or by
  `RemoveItem` down to nothing, takes its rows down. It used to leave
  them on the sheet, looking clickable, with a check list's select-all
  header still counting the old items. Items added back draw fresh.
- A `DatePicker`'s keys stop at its `DateRange`. The arrow and page
  keys used to walk the cursor onto days outside the range, where Enter
  and Space did nothing, and a calendar opened with no date started on
  today even when today lay outside. The cursor now stops on the
  range's first or last day, and the calendar opens on the nearest day
  in range. The month arrows stop at the range's first and last month
  and read muted there, where they turned into months with no day to
  pick; a live arrow now reads in full ink.
- `SliderRange` refuses a step of zero or less, which left a stepper's
  buttons and arrow keys dead or running backwards.
- Errors say where they came from. A component's errors start with its
  id ("Component 'zone': RestrictToItems applies to ComboBox only."),
  so a failure in a long build names the control. A member called on
  the wrong kind of value names the one it needs ("Text applies to a
  component; this value is an app."), where it used to say only that
  it was not valid. Reusing an id for another kind names both kinds,
  and a relative placement that names no component, or runs in a
  circle, names `RightOf` or `Below`, whichever set it; both used to
  say `Below`.
- Clicking a drop list's down pager until the last page no longer
  closes the list (#2). The pagers came and went with the page: the
  first page down added an up pager that pushed every row one row
  lower, and the last page removed the down pager, so the next click
  where the pager had been landed on an item row and picked it. A list
  longer than its window now keeps both pager rows in place at every
  page. At the first or last page the pager with nothing beyond it
  shows its arrow alone in muted ink, and a click on it does nothing.
  This applies to a `SelectBox`, `ComboBox`, and `MenuButton`.
- The accent ring around a focused control that is not a text field is
  now opt-in: `ui.FocusRings` turns it on for an app, and
  `FocusRings False` turns it off again. Without it, focus still moves
  with Tab and takes keys, and nothing marks the control. A text
  field's focused border is unchanged. (#3)
- CI's static gate runs pyvbaanalysis 2.2.1. Its Excel model now says
  `Shape.Duplicate` returns a Shape, as Excel's type library does, so
  the suppression 1.0.1 put on that line is gone.
  `requirements-dev.txt` asks for 2.2.1 or later, since 2.2.0 still
  reports the line.
- A handler a key fires runs with the control as `ReDimUI.Sender`, and
  an error it raises reaches `OnError`. An arrow on a SelectBox,
  RadioGroup, Tabs, or Stepper, Space on a CheckList, and Enter on a
  combo row or a menu command ran `OnChange` or `OnClick` with no
  sender, so a handler that read `ReDimUI.Sender` stopped at error 91,
  and any error it raised was lost without a trace.
- Enter in the command palette runs its top match. With no row
  highlighted it closed the palette, left the keys nowhere, and clicked
  the app's `DefaultButton`, so "exp" and Enter could save a form. A
  press anywhere else closes the palette, and opening it again while it
  is open keeps the control it gives the keys back to.
- Alt+Up on an open MenuButton closes it, as documented, where it ran
  the highlighted command, and a menu whose commands are all disabled
  draws its list when a key opens it. Tab in a `Confirm` with one button
  keeps focus there; the keys went back to Excel with the dialog still
  up. Focus given from code while its sheet is behind another holds
  when the sheet comes forward, where the next tick ended it.
- A combo pick fires an `OnInput` still waiting out its `DebounceMs`
  and clears a `Required` message; both waited for the next focus.
- Disabled items stay out of reach in a ComboBox: `RestrictToItems`
  committed one typed in full or begun, and Down could highlight one,
  after which Enter did nothing. An open list's highlight follows its
  item through `AddItem` and `RemoveItem`, and Enter on a row disabled
  under it picks nothing, where it could make a group header the value
  or run a disabled command.
- `Items` and `ItemsFrom` keep the pick on its item by text, as they
  keep a check list's checks; it stayed on the position, so a refreshed
  list showed another item. `ItemsFrom` reads the 2D array a range's
  `Value` gives, row by row, where it raised error 9. `Null` entries
  list nothing, where they raised error 94, and an error value lists as
  its cell shows it, `#N/A` rather than "Error 2042".
- `RemoveItem "1"` removes the item named "1"; a String that looked
  like a number was taken as a position. `AddCommand` brings back a
  menu row `RemoveItem` took, and an unknown icon name raises before
  anything is kept, where the same call with the name fixed raised
  error 5. `AccessKey` goes on the controls `Shortcut` takes and raises
  on others; on a RadioGroup it picked the first item.
- A transfer list's selection on the available side follows `AddItem`
  and `RemoveItem` and clears on `Items`, where it went on naming rows
  by number and `>` moved the wrong item. `>` and `<` read muted while
  every selected row is filtered out, Space and Enter no longer act on
  a cursor row the filter hides, a drag drops nothing while the list is
  busy or disabled, and a double-click move puts the other selected
  rows back before `OnChange` runs, so a handler that resets the lists
  keeps its work. A busy SlideBar takes no press.
- Select-all on a filtered check list with no rows showing fires
  nothing, where it fired `OnChange` for no change.
- `SetTheme` before the first `Render` only records the theme. It
  rendered the empty app, which swept the sheet's saved shapes before
  any control claimed them, drew every later builder call as it ran,
  and seeded each `WritesTo` key before the build set its value, so
  `.WritesTo("dark").Checked(True)` came up off. The Pokedex demo
  builds in that order.
- A theme change restyles everything it touches. Controls compared only
  their fill, so a theme that changed the border or the font alone kept
  the old ones, and a theme changed in place with its `With` builders
  kept every caption, hint, badge, and list row in the old colors.
  Each theme and each `With` builder now takes a new revision that the
  draw keys compare.
- `Remove` on a Label, Card, or Spinner takes its caption, hint, and
  badge, which stayed on the sheet. An Image whose shape was deleted by
  hand gets its picture back at the next draw; it came back blank.
  `Unmount False` puts away a showing tooltip and the hover look, as
  shutting down does, so neither is saved with the workbook.
- A `Toast` inside `BeginUpdate` draws with the batch at `EndUpdate`,
  where it drew the batch's changes half done. New words on a toast
  with a tone keep the ink color; they took the icon's tone color.
  `MultiLine` or `AutoGrow` set after a field first drew anchors its
  text to the top. `ContrastReport` lists two more pairings the
  controls draw: a Warning button's ink and a validation message.
- A key a text field refuses changes nothing. A letter typed over a
  selection in a `Numeric` field, or anything past a full `MaxLength`,
  deleted the selection, fired `OnInput`, and cleared the redo. Pasting
  a copied empty cell pastes nothing, where it typed a space.
- The clear button is an undo step, so Ctrl+Z brings the text back.
  Esc restores the check's verdict with the text, so a failing text
  put back shows its message again.
- A ComboBox's item follows its text when code, state, or Esc sets it,
  so its open list marks the right row and `CurrentValue` names it.
- An error value in a cell-backed field's cell reads as the cell shows
  it, `#N/A`, in `InputValue`, `WritesTo`, and the combo list's
  filter, where it read "Error 2042". Line breaks written from code or
  state become the field's own, so Up, Down, Home, and End count them;
  End in an empty field, and Up or Down in an empty multi-line one,
  raised error 9 unseen. `Masked` raises on a ComboBox, as documented.

## 1.0.2 - 2026-09-23

- A toast at the top of its tray that runs taller than the room below
  the rail lifts the rail, so the whole card stays in the window. The
  rail was placed for a one-line toast, so near the window's bottom a
  three-line toast ended below it. A `ToastTray` anchor still holds the
  tray exactly where it is pinned.
- A toast shown while its sheet is behind another fits to the window
  once the sheet is in front. Off its sheet the window cannot be read,
  so the card was fitted to A1:P36 instead, and a `MaxWidth` toast could
  come out wider than the window had room for, and stayed so.
- A calendar highlight deleted by hand, and drawn again, goes back
  behind the week's digits. It came back on top of them and hid the day
  under it until the calendar closed.
- A tooltip showing when its words change, as when its control is
  disabled under the pointer and has a `DisabledReason`, shows the new
  words. It kept the old ones until the pointer left and came back.
- Requires ROneCOne 1.9.1 or later, and is tested against 1.10.1.

## 1.0.1 - 2026-09-23

- A DatePicker's calendar opened by a click now closes on a press off
  it or a move of the grid selection, as every drop list does. Only a
  picker holding keyboard focus closed that way, and a click never
  gives a picker focus, so a calendar opened with the mouse stayed open
  until another control was clicked. `TestListDismiss` now opens one by
  a click and dismisses it both ways.
- One press off a focused SelectBox or DatePicker closes its list and
  ends its focus; it took two. And a press that closed a list no longer
  ends that control's focus later, when Tab or `Focus` gives it focus.
  The watch for an open list and the watch for keyboard focus shared one
  pending press, so the list's watch used up the press the focus needed,
  and a press it recorded outlived the list. Each keeps its own now.
  `TestListDismiss` presses and releases apart, as a real click does.
- The calendar opens in about a fifth of the time: 13 ms instead of 73
  in the bench, 4 instead of 12 to pick a day and close, and 1.5
  instead of 6 to turn a month. A session's first opening, which also
  pays for loading the code, takes 65 ms instead of 110. The calendar
  drew a shape per day, 47 shapes in all. A week is now one line of
  text with a center tab stop over each day column, and the date held,
  today, the key cursor, and the day under the pointer each get one
  mark shape, 14 shapes in all. The digits now sit exactly under the
  weekday initials, which already used tab stops. A click on a week
  picks the day under the pointer.
- Drop lists open faster. A list draws its first row in full and
  copies it (`Shape.Duplicate`) for the others, since a copy carries the
  size, border, margins, gutter tab stop, font, and click. A SelectBox
  over 2,000 items opens in 5.7 ms instead of 14.5, a ComboBox in 7.9
  instead of 16, and a three-command MenuButton in 3.5 instead of 5.5.
  Drawing-object protection refuses a copy even under
  `UserInterfaceOnly`, so on a surface `ProtectSurface` protected, the
  protection lifts for each copy and returns as it was, selection rule
  included, about a tenth of a millisecond each time. A sheet its host
  protected draws every row in full, as before.
- Opening or closing a drop list or calendar now runs with screen
  updating off, so Excel paints the list once instead of part by part.
  Before, only a flush that repainted several controls did.
- A part deleted when a list or calendar closes is known to be gone,
  so drawing it at the next opening skips a lookup that could only
  miss.
- Tooltips stay out of the way. A tip waits for the pointer to hold
  still on its control for the tooltip delay, as Windows does, instead
  of counting from the moment the pointer arrived, so a pointer passing
  over a control no longer raises one. On a control no taller than 48
  points the tip sits clear of the control, below its hint or message;
  it used to sit 18 points under the pointer, over the control's lower
  half and its hint. On a taller control it still sits under the
  pointer, and it goes as soon as the pointer reaches it, so a click
  there lands on the control instead of dismissing the tip. An open drop
  list or calendar shows no tip over its rows, where the owner's tip
  used to appear after a rest on a row. A tip is sized by Office to fit
  its words, replacing an estimate from the character count that could
  cut a long tip short. With no room below the pointer, the tip sits as
  far above it as it would below, so a small move up keeps it.
- A tooltip showing on a Label, Card, Spinner, or Overlay goes when the
  control is removed, and a tooltip showing at `Shutdown` goes too. Both
  stayed on the sheet over whatever lay under them, and a workbook
  saved then kept the tip.
- A MenuButton's menu runs as wide as its widest command needs. It took
  the button's width, so a longer command ran past the menu's edge. A
  menu that would pass the window's right edge ends at the button's
  right edge instead, and a press anywhere on the widened rows counts as
  a press on the menu. A command paged into view widens the open menu
  when it needs more room, and an open menu follows its button when the
  button moves.
- A click on a toast puts it away at once. The close button, the action
  button, and the card only set the toast to expire, and while the
  pointer rested on the toast (as it does for any click) the rest held
  the countdown and undid the dismissal, so the close button seemed to
  do nothing. The close button now runs no handler, and a toast already
  on its way out takes no more clicks, so a double-click on its action
  runs the handler once.
- `Shortcut "^s"` gives a control a keyboard shortcut: the key clicks it
  exactly as a mouse click does, while its sheet is in front. It needs
  Ctrl or Alt, or is a function key from F1 to F15, so it never takes a
  key a field types; Ctrl+Alt takes only a named key such as Del, since
  Windows sends AltGr as Ctrl+Alt. A code OnKey would refuse is refused
  at once. The tooltip and alternative text name it ("Saves the form
  (Ctrl+S)"). It goes on a control a click acts on as a whole, a label
  or card only with a click handler, and a slider, list, or table
  raises an error. A disabled or hidden control passes the key to the
  next that declares it. A field that took a shortcut's key for editing
  gives it back when it lets go, and any other focused control lets the
  key through. Removing the control gives the key back. The keys ReDim
  binds are also listed in a hidden workbook name, so a reset of the VBA
  project cannot leave one bound past the workbook's close; a key left
  bound by a reset goes back to Excel on its first press. Access keys
  get the same care. `ReDimHost.bas` gains `RdxShortcut` and
  `RdxCapturedCodes`, so import both files again.
- `OnClick` on a toast runs when the card itself is clicked, with the
  toast as `ReDimUI.Sender`, before the toast goes.
- `ActionBorder rgb` colors a toast's action button border, the theme's
  accent until set.
- A toast is as tall as its message. It was a fixed 240 by 40 points,
  so a message longer than a line spilled out of the card and over the
  toast below. The card now wraps the message and grows to fit it, up
  to 160 points, where the message ends in an ellipsis; `MaxHeight`
  moves that cap. `MinWidth` and `MaxWidth` let the card fit its words
  between them, 240 until set. Toasts stack by their heights, and the
  corner radius holds at 8 points however tall a card grows. A limit set
  after the toast drew, as `app.Toast(message).MaxHeight(300)` does,
  brings back the words the default cap cut. A message with no space to
  cut at, such as a long path, ends in an ellipsis inside the card.
- The runtime no longer declares what it never reads: the constants
  `ROLE_UNSET` and `KIND_NONE` and the fields `mHasRect`,
  `mHoverSinceMs`, and `mAppliedValue`, which code inspections listed
  as unused. A click reads its part once instead of at every branch.
  The Widget Gallery and the pump and timer tests await their delays as
  statements rather than into a variable they never read.
- A session's first build runs up to about 20 ms longer than with
  1.0.0, from the runtime's size: VBA takes longer to run a class's
  code the first time the more code the class holds, and adding as many
  never-run lines to 1.0.0 costs it the same. Rebuilds and warm builds
  are unchanged. An app with no shortcuts and no drop-downs skips their
  bookkeeping at each render and flush.
- Requires ROneCOne 1.9.1 or later, and is tested against 1.10.1.
- CI's static gate runs pyvbaanalysis 2.2.0, which also fails on a
  variable or constant that is declared but never used or never read,
  and analyzes against the ROneCOne release ReDim ships instead of
  ROneCOne's main. One line carries a suppression: the analyzer's Excel
  model says `Shape.Duplicate` returns a ShapeRange, where Excel's type
  library says Shape.

## 1.0.0 - 2026-09-23

- ReDim no longer recases the host project's identifiers. VBA keeps
  one spelling per name across a project, and a declaration in any
  module can set it. ReDim's lowercase parameters and locals (`value`,
  `color`, `name`, `caption`, `source`, `size`, `title`, `target`, and
  more) turned `.Value`, `.Color`, `.Name`, `.Caption`, and ten other
  Excel members lowercase in every host module, then into every
  exported diff. Its camelCase parameters did the same to its own
  `AppId`, `ComponentId`, and `OwnerApp`. Every such name is renamed,
  so the runtime declares nothing in a casing that differs from the
  Excel, Office, VBA, or stdole type libraries or from its own members.
  A new `tests/python/test_casing.py` holds the runtime and the demos
  to that rule and exports ROneCOne, ReDim, every demo, and sample host
  code through the VBE as one project, requiring every token back as
  written. Against 0.19.2 the same export recased 67 tokens across the
  modules, 20 of them in the host code alone.
- The samples name their app variable `ui`. A variable named `app`, as
  the README and every demo wrote it, collides with `ReDimUI.App`, so
  exporting a demo rewrote its own calls as `ReDimUI.app(...)`. The
  demos also drop a dozen locals that clashed the same way, such as
  `key`, `doc`, `names`, `typeName`, and `startRow`. The test modules
  keep `app`; they are never exported.
- Requires ROneCOne 1.9.1 or later, and is tested against 1.10.0.
  Earlier ROneCOne releases declared `value`, `text`, `cells`, and
  similar names in lowercase, so importing one recased `.Value`,
  `.Text`, `.Cells`, and more across the whole project. ROneCOne 1.9.1
  fixed that in its
  [issue #6](https://github.com/WilliamSmithEdward/ROneCOne/issues/6),
  and 1.10.0 changes only its demo workbooks. With ROneCOne in the
  round trip, two internal ReDim names turned out to clash with
  ROneCOne's own, `InstancePointer` and `FieldText`. They are now
  `FactoryPointer` and `FieldTextNow`.
- Five ROneCOne locals share a name with a ReDim public member: `at`,
  `currentValue`, `itemCount`, `items`, and `steps`. Imported in the
  README's order, ReDim's spelling wins and only ROneCOne's own module
  changes. Importing a newer ROneCOne into a project that already
  holds ReDim reverses that: `.At`, `.CurrentValue`, `.ItemCount`,
  `.Items`, and `.Steps` come back lowercase in ReDim and in every
  module that calls them. Reimporting ReDim afterward restores them.
- Breaking for named-argument callers only; positional calls are
  untouched. Public parameters renamed: component factories and
  `Component` (`componentId` to `targetComponentId`), `Mount`, `App`,
  `HasApp`, `Navigate`, and `NavigatesTo` (`appId` to `targetAppId`),
  every handler registration (`procName` to `handlerProc`), `Text` and
  `BusyText` (`displayText`), `Visible`, `Enabled`, `Checked`, and
  `Bold` (`flag`), `Fill` (`fillRgb`), `TextColor` (`inkRgb`), `At` and
  `ToastTray` (`rangeAddress`), `Toast` and `Confirm` (`messageText`,
  `titleText`), `WindowTitle` (`titleText`), `ProtectSurface`
  (`protectOn`), `AutoPump` (`pumpOn`), `PinPumpCursor` (`pinOn`),
  `MultiLine` (`multiLineOn`), `WithSelectAll` (`selectAllOn`),
  `ItemsFrom`, `ChosenFrom`, and `CheckedFrom` (`itemSource`),
  `ItemTextAt`, `ChosenTextAt`, `SetItemChecked`, and `IsItemChecked`
  (`itemNumber`), `SliderRange` (`lowestValue`, `highestValue`),
  `FontSize` (`fontPoints`), `Source` (`picturePath`), `PacedMs`
  (`paceMs`), `Value` (`newNumber`), `InputValue` (`assignedText`),
  `WithFont` (`fontFace`, `fontPoints`), and `WithPrimary`
  (`primaryRgb`, `onPrimaryRgb`).
- Clicking a cell-backed `TextInput` selects its cell. The frame drawn
  over the anchor cell carries the dispatcher, so the click reached
  ReDim instead of the grid and the cell never took focus; the click
  also ran the input's `OnChange` handler as if an edit had happened.
  It now selects the cell, where typing replaces the value and F2 edits
  it in place, and fires nothing. A cell-backed `ComboBox` face gets the
  same fix: its click selects the cell as well as opening the list.
- `SelectBox` drop lists window. A select drew every item as a row, so
  a long list ran down the sheet. It now shares the combo's windowed
  list: eight rows by default, clickable pager rows at the edges with
  the count beyond them, and the selection scrolled into view on open.
  The new `ListRows(n)` sets the window for either control. An open
  select also claims only its drawn depth from the slider press watch,
  where it used to claim the height of every item.
- A combo reopens onto its whole list after a pick. The list filters by
  the field text, and after a pick the text is the item, so reopening
  showed that one item until the user backspaced it away a character
  at a time. While the text names an item and the user has not edited
  it since the list opened, the list is now unfiltered and scrolled to
  that item, and Down walks on from it. The first edit filters as
  before.
- Keyboard picks fire `OnChange`. Enter on a highlighted combo option
  wrote the value and the state but never ran the handler, so ReDex's
  species combo, whose handler fetches the pick, ignored every keyboard
  pick. A pick now commits the way the field's other commit paths do,
  from the keyboard or the mouse: when it changes the value it writes
  the `WritesTo` state and runs the handlers once, and a re-pick of the
  value already there only closes the list. Mouse re-picks used to fire
  anyway, and a list that now opens on its current item makes clicking
  that row the natural way to close it. The float combo measures the
  change against the text it held when focus arrived, the same
  reference its Enter and Tab commits use. The `SelectBox` follows the
  same rule, as `RadioGroup` already did for its selected row.
- The API guide documents what was discoverable only from source: a
  `SelectBox` shows its `Text` as the placeholder while nothing is
  selected, and `Value(0)` clears the selection back to it.
- Every source a release ships, both runtime files and each demo
  module, opens with the MIT license text, the repository link, and a
  version line, `ReDim 1.0.0 (2026-09-23)` for this release, so a
  module copied out of a release or a project still says what it is
  and under what terms. `tools/stamp_release.py` writes the header from
  `REDIM_VERSION`, the CHANGELOG date, and `LICENSE`, and a source
  guard fails the suite when any file's header falls behind.
- Releases attach `ReDimUI.cls` and `ReDimHost.bas` beside the demo
  workbooks, and v0.19.2 carries them now too. The README installs from
  those copies: the repository stores the sources with LF line endings,
  and the VBE imports an LF-only class file as a standard module with
  its header pasted in as code. The release copies are CRLF.
- Smaller fixes on the same paths: an open drop list repaints its rows
  when the theme changes, and removing a combo or select deletes its
  pager rows instead of leaving them for the next render to sweep.
- Lists redraw only the rows that changed. Each row keeps the state it
  was last drawn with, and a row whose place and font held rewrites only
  its fill, text, and ink. Against 0.19.2, as the fastest of three warm
  runs on one machine with both versions benched back to back: toggling
  one row of a 150-row `CheckList` fell from 351 ms to 0.9 ms and its
  select-all from 358 ms to 40 ms. A `TransferList` row click among
  1,000 items fell from 22 ms to 0.4 ms and a page from 28 ms to 2.5 ms.
  An arrow key in an open 2,000-item combo fell from 10 ms to 2.9 ms.
  Opening a 2,000-item `SelectBox` took 8.4 seconds when it drew every
  row and takes 15 ms windowed, and a pick from it fell from 577 ms to
  3.6 ms. Paging it costs 11 ms, where 0.19.2 only scrolled rows it had
  already drawn.
- Tick boxes, radio groups, steppers, toggles, sliders, and drop-list
  carets skip their part writes when nothing they draw has changed.
  `Render` still rewrites them, so a part moved or deleted by hand comes
  back.
- Item text is read from a snapshot kept beside the items. A VBA
  `Collection` read by index walks from its start, so filtering,
  matching, and drawing a long list paid a walk per read. A closed list
  also no longer rebuilds its matches on every change: adding or
  removing an item at the front of a closed 2,000-item select fell from
  about 0.22 ms to 0.05 ms. `ItemsFrom` reads a range one area at a
  time instead of one cell at a time, and loading 2,000 items fell from
  3.3 ms to 0.5 ms.
- The frame pump visits only the controls that tick: spinners, toasts,
  sliders, the text fields and combos whose caret blinks, and open drop
  lists. An idle frame over 100 controls fell from 0.27 ms to 0.007 ms.
  Focus moving from one text field to another keeps the typing keys
  bound instead of releasing and binding them again.
- `Unmount` deletes an app's shapes in one pass over the sheet and one
  batch delete. It used to try every name a part could have, eight per
  control and four more per item, and each absent part cost a failed
  lookup. Unmounting 100 controls fell from 87 ms to 5 ms.
- Rebuilding a surface over its existing shapes takes about half as
  long: ReDex 95 ms to 52 ms. A session's first build runs 70 to 130 ms
  slower than in 0.19.2 (Mission Control 152 ms to 233 ms, ReDex
  439 ms to 565 ms). That cost is paid once per session: the same build
  repeated from empty sheets costs what it did in 0.19.2 or less, ReDex
  229 ms to 184 ms. The Widget Gallery builds its new section as well,
  so its numbers do not compare; it now batches its build between
  `BeginUpdate` and `EndUpdate`, and rebuilds in 66 ms against 0.19.2's
  103 ms.
- `ReDimUI.Shutdown` forgets window registrations and the back stack
  along with the apps. A multi-window surface rebuilt after a
  `Shutdown` in the same session failed in its `NavBar` with "No
  mounted app is named", because the bar still tabbed to the unmounted
  windows. The new bench's warm builds found it in Navigator and ReDex.
- `tools/bench.py` times framework scenarios, the controls new in
  1.0.0, and each demo's first build, warm build, and rebuild in a
  live Excel, and compares against a saved run with `--save NAME` and
  `--compare NAME`. The numbers in these notes come from it.
- Drop lists dismiss the way native ones do. A press anywhere off an
  open list's face and rows closes it, and so does a move of the grid
  selection, a click on another control, or another list opening, so
  one list is open per app. The pump watches an open list for this the
  way it already watched a focused field.
- The current item's row in a drop list, and each selected row in a
  transfer panel, shows a check in a gutter every row shares, so a
  selection no longer reads by color alone.
- An open list with nothing to show says so in an inert row: "No
  matches" under a combo's filter, "No items" in an empty select. The
  list used to vanish while it was still open.
- A drop list that would run past the bottom of the visible window
  opens upward when there is more room above its face.
- Transfer paging arrows grow from 14 to 18 points square, and the new
  toast close button matches: 24 pixels at 96 DPI, the WCAG 2.5.8
  minimum target size.
- The insertion bar blinks at the Windows caret rate from
  `GetCaretBlinkTime` instead of a fixed 500 ms, and stays solid when
  Windows is set not to blink.
- Reduced motion: when the Windows "Show animations" setting is off,
  toasts appear, move up, and leave without sliding or fading.
  `ReDimUI.ReduceMotion` overrides the setting, and
  `ReDimUI.MotionReduced` reports the result.
- Toasts gain a close button, a reading-time TTL, tones, and actions.
  Without a TTL a toast stays three seconds plus 60 ms a character,
  from four seconds to twelve; it was a flat three. `Primary`,
  `Success`, `Warning`, and `Danger` give it an info, success, warning,
  or error tone, an icon and a matching edge. `.Action "Undo",
  "Module.Proc"` adds a button that dismisses the toast and runs the
  handler, and gives the toast four more seconds. The new parts cost
  time: a toast appears in 4.1 ms instead of 1.9 ms, and a frame that
  slides five toasts takes 1.0 ms instead of 0.45 ms.
- A `Warning` style variant joins the set: amber, from
  `theme.WarningColor`, which picks a dark shade on light surfaces and
  a bright one on dark surfaces so it keeps its contrast in every
  theme.
- Every control's shape carries alternative text for screen readers,
  written from its kind, text, and state ("Save, button, unavailable",
  "Agree, checkbox, checked"), and `AltText` replaces it. Labels, cards,
  and toasts carry none, since their text is what a reader announces,
  and progress reads in 5 percent steps. Both keep the writes off the
  busiest update paths.
- `ReDimUI.ThemeHighContrast` is white and yellow on black, and every
  pairing the controls draw passes WCAG AA. `theme.ContrastReport`
  lists those pairings for any theme with their ratios and verdicts,
  and `ReDimUI.ContrastRatio` computes one pair. The light and dark
  presets keep their look; each falls short in one place, the field
  edge, at 1.32:1 and 2.01:1 against the 3:1 non-text minimum.
- A switched-on toggle's knob takes the theme's `OnPrimary` ink instead
  of fixed white, as Windows draws it: unchanged on the light theme,
  dark on the dark theme's green, and black on the high-contrast
  yellow, where white would vanish.
- Every control takes keyboard focus. Tab and Shift+Tab walk an app's
  controls in creation order, positive `TabIndex` values first, and
  `TabIndex(-1)` leaves a control to `Focus` alone; `ui.FocusFirst` and
  `component.Focus` place focus from code. A click still focuses only
  text fields, so the grid keeps its keys after a button click. A
  focused control wears an accent ring, focus that lands off screen
  scrolls into view without moving the selection, and Esc, a click
  elsewhere, or a moved selection ends it.
- Keys for each control: Space and Enter click a button, and Space
  toggles a toggle or tick box. Arrows move a radio group's selection
  and step a stepper or slider, with Page Up, Page Down, Home, and End
  for bigger moves. A check list and a transfer list show a row cursor
  that Space toggles, and Enter moves a transfer panel's selection
  across.
- A focused `SelectBox` works like a native drop-down. Arrows change
  the selection while it is closed; Space, Alt+Down, or F4 opens it on
  the current item; Enter or Space takes the highlight and Esc closes;
  typed letters jump to the next item that starts with them.
- `Confirm` takes keyboard focus: OK holds it, Tab stays among the
  dialog's buttons, Enter confirms, Esc cancels, and closing the dialog
  returns focus to where it was.
- `ui.DefaultButton "save"` names the button Enter clicks, from a text
  field after its commit and from any control that does not use Enter
  itself.
- `AccessKey "s"` underlines the letter in a button's or tick box's
  text and clicks the control on Alt+S while its sheet is in front. The
  chords are bound only while such a sheet is active.
- Behavior changes for float fields. Tab and Shift+Tab commit and move
  focus to the next control; they used to commit and leave. A combo's
  Esc closes an open list first and then reverts; it used to clear the
  value. A click on another control commits the field before that
  control's handler runs; it used to depend on the pump catching the
  press. `Clearable` adds a clear button to a float `TextInput` or
  `ComboBox`, for the clearing Esc no longer does.
- The captured keys add Page Up, Page Down, Shift+Tab, Alt+Down, Alt+Up,
  and F4, bound and released from one list the two share.
- Float fields edit like a Windows text box. Shift with the arrows,
  Home, and End selects, and the selection shows in the accent color.
  Ctrl+Left and Ctrl+Right move by words, Ctrl+Backspace and Ctrl+Del
  delete them, Ctrl+Home and Ctrl+End reach the ends of the text, and
  Ctrl+A selects everything. Typing, a delete, or a paste replaces the
  selection. Fields had no selection before.
- Ctrl+C, Ctrl+X, and Ctrl+V copy, cut, and paste through the Windows
  clipboard as Unicode text. A single-line field turns pasted line
  breaks and tabs into spaces, and drops the line break a copied cell
  brings along.
- Ctrl+Z undoes and Ctrl+Y or Ctrl+Shift+Z redoes, up to 100 steps
  since the field took focus. Typing groups into one step until it
  pauses for a second.
- Punctuation and symbols type into fields: `@`, `#`, brackets, quotes,
  slashes, and the rest, bound by character so each follows the active
  keyboard layout. Fields took only letters, digits, space, minus,
  period, and comma, so an email address could not be typed.
- A click puts the caret where it lands, the click that focuses the
  field included, and a double click selects the word under the
  pointer. The caret used to go to the end. A click on a focused
  combo's text now places the caret instead of toggling the list; the
  arrow at its right edge still toggles it.
- `CurrentText` on a focused float field returns the text. It carried
  the insertion bar, or the space the bar blinks to.
- New builders for float fields. `Placeholder` shows a muted hint while
  the field is empty. `OnInput` runs a handler after every edit, or once
  typing pauses for `DebounceMs`. `Numeric` keeps digits, one decimal
  separator in the locale's form, and a leading minus. `MaxLength` caps
  the text and counts it under the field. `Validates` checks the text on
  commit with a function that returns a message: an invalid field shows
  the message and a danger border and rechecks on every edit until the
  text passes, and `ValidationError` reads the message.
- `AutoGrow(maxLines)` lets a multi-line `TextInput` grow with its text
  from the height it was given, up to six lines by default, and shrink
  back, moving anything placed `Below` it.
- Combo lists bold what the typed text matched in each row, and the
  first item the text begins shows the rest of its name after the caret
  in muted ink. Right at the end of the text or Tab takes it, in the
  item's spelling. This changes Tab, which committed exactly what was
  typed. `RestrictToItems` makes a float combo commit only an item: the
  one its text names or begins, or an empty text, and otherwise the
  text it had when focus arrived.
- The editing costs a little time. Against 0.19.2, a key in a float
  field takes 0.11 ms instead of 0.07 ms, and focusing a field from the
  grid binds 139 keys instead of 86 and takes 1.3 ms instead of 0.8 ms.
  A key in an open 300-item combo takes 0.95 ms instead of 1.9 ms:
  bolding the matches gives back part of what the item snapshot saved.
- `ui.PointerEffects` gives an app hover and press looks: the fill under
  the pointer moves toward its ink, 8 percent on hover and 16 while the
  press that went down on it is held, and a check box's or radio
  button's edge takes the accent. An open drop list's highlight follows
  the pointer as it moves. Off by default, since the pump then reads the
  pointer every frame the app's sheet is in front.
- `Tooltip` shows a note once the pointer rests on a control for the
  Windows tooltip delay, and `DisabledReason` says why a disabled
  control is disabled. A press puts the tooltip away until the pointer
  leaves, and screen readers get the note with the control's
  alternative text.
- Held presses repeat: a stepper's minus or plus, a drop list's pager,
  and a transfer panel's paging arrow repeat after the Windows keyboard
  delay, at the keyboard repeat rate. A held stepper fires `OnChange`
  once at the release, as a slider drag does.
- Transfer lists take gestures. A double click on a row moves it
  across. A press dragged along a panel selects the rows it covers, and
  dragged onto the other panel it outlines that panel and moves the row,
  or its whole selection, on release. Clicking a selected row again to
  deselect it now has to wait out the double-click time.
- A slider shows its value in a bubble over the thumb while a drag or
  keyboard focus moves it, and a toast under the pointer stops counting
  down until the pointer leaves.
- A visible stepper or transfer list on the active sheet keeps the pump
  armed for its press watch, as a slider does, at the cost of a
  key-state poll each frame.
- A sheet coming back to the front arms the pump again. The pump stops
  itself once no app has work, and leaving a sheet takes a slider's
  press watch with it, so a slider on a sheet the user left and came
  back to took taps but no drags until something else started the pump.
- A slider's press watch acts only while its sheet is in front. A press
  on the active sheet could start a drag on a slider in another sheet
  whose track sat at the same position.
- Copy, cut, and paste retry briefly when another program holds the
  clipboard open, instead of doing nothing.
- Items can carry values apart from their text: `AddItem "Medium", ,
  20` or `ItemsFrom texts, values`. A pick writes the value to
  `WritesTo`, with its own type, instead of the text; check lists and
  transfer lists join the values of what is checked or chosen, and
  `ItemValueAt` and `ChosenValueAt` read them. A value belongs to its
  item's text, so it follows the item between transfer panels.
- A `SelectBox` can disable items and group them. `ItemEnabled n, False`
  leaves an item showing, muted, but out of reach of clicks and keys.
  `AddGroup "Fruit"` adds a bold, muted header row that labels the items
  after it and is never picked; it takes a place in the item numbers.
- `Reorderable` gives a transfer list up and down arrows in the chosen
  panel's header, and Alt+Up and Alt+Down, which move the chosen side's
  selected rows a place as a block. The chosen order is part of the
  value, so a move writes `WritesTo` and fires `OnChange`.
- Typing while a transfer list or check list has the keys filters it.
  A transfer panel's header shows the filter and how many rows match; a
  check list hides the rows that do not and its select-all works on the
  rest; Backspace takes a letter off and Esc clears the filter. Moves
  under a filter take only the rows that show, and hidden rows keep
  their selection and their checks.
- Controls take a `Caption` above them and a `Hint` below them, in
  small text. A float field's validation message takes the hint's
  place while it shows, and a `MaxLength` count keeps its corner.
- `Required` marks a field: its caption ends in an asterisk in the
  danger color, and an empty commit shows "Required", or the message
  given, as a validation message before any `Validates` check runs.
  `ErrorText` shows an error from outside the field, such as a server's
  answer, with the same danger border until it is set to "".
- `Skeleton` is a loading placeholder: a muted block, or with
  `SkeletonLines n` the bars of the text it stands in for, the last one
  short. It pulses gently unless motion is reduced; hide it and show
  the content when the data arrives.
- `Tabs` is a tab strip: one tab per item, each as wide as its text,
  the tab shown in bold over an accent bar. A click, or Left, Right,
  Home, and End, shows a tab, writes it to `WritesTo`, and fires
  `OnChange`. Controls join a tab's panel with `OnTab "tabs", 2` and
  show only while their tab does; their own `Visible` still applies on
  the tab, and a change made while the tab hides them waits for it. A
  panel that hides closes its open list and commits a focused field,
  and hiding the strip hides every panel.
- `DatePicker` is a date field with a month calendar that drops under
  its face. The face shows the date in the system's short date or a
  `DateFormat` pattern; a day click, or the arrows, Page Up, Page Down,
  and Enter with the keys, picks, writes a `Date` to `WritesTo`, and
  fires `OnChange`. `DateRange` limits the days that pick, `PickDate`
  and `PickedDate` set and read the date from code, and the calendar
  closes the way a drop list does.
- `Table` is a data table: `Columns` and `AddRow`, or `TableFrom` a
  Range or two-dimensional array. Number columns align right, long
  cells end in an ellipsis, and `ColumnWidths` and `ColumnFormat` set
  widths and `Format$` patterns. A header click sorts by its column,
  ascending and then descending, stably, with empty cells last. A row
  click selects, writes the row's first cell to `WritesTo`, and fires
  `OnChange`; the keys move the selection, and rows that outgrow the
  table page behind a footer.
- The new controls, as the fastest of three warm runs: a tab switch
  that hides ten controls and shows ten takes 4 ms. A date picker's
  calendar opens in 78 ms, drawing its card and 42 days, turns a month
  in 6 ms, and picks and closes in 13 ms. A 1,000-row `Table` fills
  from an array in 34 ms, sorts in 7 ms, pages in 5 ms, and picks a row
  in 0.6 ms. A `Skeleton` frame costs 0.6 ms. Calendar days and table
  rows whose place and font held rewrite only their text, fill, and
  ink, which took a month turn from 45 ms to 6 ms and a page of rows
  from 16 ms to 5 ms.
- The Widget Gallery shows the new controls in a tabbed section beside
  the others: a form with captions, a hint, a required field with a
  check, a date picker, and a grouped select; a sortable table that
  pages; and a skeleton that a button swaps for its content. It turns
  on `PointerEffects`.
- `Icon("Save")` draws a Windows icon on a button or label, before its
  text or alone. There are 71 named icons, drawn in Segoe Fluent Icons,
  or Segoe MDL2 Assets where Windows lacks it. An icon-only button
  reads to screen readers as its icon's name. `ReDimUI.IconGlyph` and
  `IconFont` give the character and the font for text of your own.
- `ReDimUI.ThemeSystem` takes the Windows look: the dark theme under
  dark mode, and the Windows accent color as the primary, moved until
  it reads against the surface and under its ink. `ui.FollowSystemTheme`
  keeps an app on it. Switching Windows to dark mode or to another
  accent re-themes the app the next time a sheet activates, the
  selection moves, or the pump runs.
- `Badge` is a pill for a count or a status word that widens to its
  text. `BadgeText` puts one on any control's top-right corner, such as
  the unread count on an Inbox button.
- `Stack` lays out the controls that join it with `InStack`: down or
  `Across`, a `Gap` apart, inside a `Padding`, with `Stretch` to give
  them its width. It counts a caption's row above a field and the note
  row under one, and grows to its content. A member that hides,
  grows, or gains a caption moves the members after it once the flush
  ends, and only members whose place changed redraw. Stacks nest, and
  a hidden stack hides its members.
- `Expander` is a collapsible section: a header with a chevron, and a
  panel of controls that join it with `InExpander`, shown while it is
  open. In a stack, what follows moves down as it opens and back up as
  it closes.
- `MenuButton` drops a menu of commands added with `AddCommand`, each
  with its handler and an icon in the gutter. A pick runs the command
  with the menu as the sender and the face keeps its text; group
  headers, disabled commands, windowing, and the keys come from the
  drop list.
- `ui.CommandPalette` adds a searchable palette of everything the app
  can do, opened with Ctrl+Shift+P or `OpenCommandPalette`: the app's
  own `AddCommand` entries, every button with a click handler, and
  every menu command. Typing filters it, a pick runs the entry, and
  leaving it runs nothing.
- A `Table` filters as you type while it has the keys: a row stays when
  any cell holds the text, the footer names the filter, Backspace takes
  a letter off, and Esc clears it. `FilterRows` filters from code and
  `ShownRowCount` counts what passes. Ctrl+C copies the selected row, or
  the rows shown, under the header as text that pastes into cells with
  dates intact, and `ExportTo` writes the rows shown to a range in one
  write. `EmptyText` words an empty table, a filter that matches nothing
  says so, and `OnRowOpen` runs on a row's double click or on Enter.
- `Masked` turns a float `TextInput` into a password box: dots on the
  face, nothing for copy or cut, and "Edit field, masked" for screen
  readers, while `InputValue` and `WritesTo` keep the text.
- `ui.ValidateAll` checks a form before a save: every enabled float
  field's `Required` and `Validates`, run as a commit runs them, on the
  panel shown or not. Each failure shows its message, the first in Tab
  order takes focus with its tab turned to or its expander opened, and
  the result says whether everything passed.
- `ReDimUI.SetUIText` translates the words ReDim draws or announces on
  its own: empty rows, footers, list headers, the palette's hint, the
  default busy, required, OK, and Cancel texts, and every screen-reader
  description. Fifty-four keys, listed by `UITextKeys`, take `{0}`-style
  fills in any order, and `ResetUIText` restores English.
- A switch's knob glides across when it flips, and a tab strip's bar
  glides to the tab shown, settling in about 200 ms at any frame rate.
  Both land at once under reduced motion, and a glide caught mid-way by
  another change carries on from where it is.
- `Sparkline` draws a small trend line through numbers from an array, a
  Collection, a Range, or a ROneCOne sequence, with a dot on the last
  value, in its variant's color, and a spoken summary of the count, low,
  high, and last value. It redraws only when its values, place, or color
  change.
- A click on an open drop list's pager could close the list instead
  of paging it. The watch that closes a list, or ends a field's focus,
  on a press off the control decided at the press, from a pointer
  reading, before Excel delivered the click on the pager. A press
  judged off now takes effect once the button is back up and a short
  grace has passed, and a click on the control in between cancels it.
- Performance, measured back to back against the build before these
  additions. First draws on a sheet that held none of an app's shapes
  skip the name lookups that could only miss: a missed Shapes(name)
  scans the whole sheet and raises, 0.1 ms at 100 shapes and 0.8 ms at
  1,000. A 100-control first render drops from 180 to 168 ms and a date
  picker opens in 74 ms instead of 80. Every drawn control used to call
  the tab-panel, caption, badge, focus-ring, and stack steps whether or
  not it used them; VBA pays for a procedure's first call in a session,
  so each call now waits on its feature. The UI text defaults moved from
  a 54-case function to one constant. Animation checks ask one kind at a
  time, where one expression asked Windows about motion for every
  ticking control each frame, and the Windows motion setting is read at
  most once a second. Switches and tab strips join the frame loop only
  while they glide. The first render of a small form, cold, fell from
  134 ms to 116 ms of the 99 it took before these additions.
- The numeric keypad types into a focused field. Its digits and
  operators were never captured, so they went past the field to the
  grid, where a protected sheet refused them with a warning and an
  unprotected one started editing the active cell. The decimal key
  types the separator Excel uses.
- A transfer list's hover and selection fills, the keyboard cursor on
  a list row, a tab's hover fill, and an expander's header round their
  corners, 4 points, inside their rounded panels. Rectangles left by an
  older build round on their next full draw.
- A `Toggle`, `TickBox`, or `Expander` bound with `BindValue` follows
  `True` and `False` in the store. `BindValue` set a number these
  controls never read, so a switch bound to a state that code changed
  stayed where the user left it.
- A new demo, `ReDim_Expense_Tracker.xlsm` (`BuildExpenseTracker`), is
  a small bookkeeping app built from the 1.0 controls: a stacked form
  checked by `ValidateAll`, a table that filters as you type and
  deletes behind a confirm with an Undo toast, a sparkline and a badge
  over six months, a budget meter, an Actions menu, the command
  palette, and a dark mode switch bound both ways. It keeps its
  expenses on a hidden sheet and seeds six months of samples on first
  open.
- A source guard fails on a statement whose last call wraps several
  arguments in parentheses, `ui.Card("c").AtRect(0, 0, 9, 9)` for one,
  which only a compile in Excel reported before.
- `ui.StateKeys` lists the state store's keys in the order first set,
  so a host can save the whole store without keeping its own list.
- A guard test fails when a public member ships without a mention in
  the guides, and the members it found undocumented now have one: the
  theme tokens, the app readers, `FlushDirty`, `ReDimUI.Version`,
  `OnCancel`, and `JobOnCancel`.

## 0.19.2 - 2026-08-02

- The item-source readers no longer trip over an object. `ItemsFrom`,
  `ChosenFrom`, and `CheckedFrom` reached a bare `IsArray(source)` when
  the source was neither a Range nor a Collection, and `IsArray` on a
  ROneCOne dereferences its default member, which raises on the zero
  arguments the dereference supplies. Under the pump's active error
  handling that raise was swallowed, leaving `Err` dirty for the next
  reader, and a ROneCOne sequence fell through every branch and yielded
  no items at all. The three readers now share one `ReadItemSource`
  helper whose array test runs through a `IsObject` guard, so `IsArray`
  appears exactly once in the runtime and never receives an object.
- Item-source readers accept a ROneCOne sequence. A `SelectBox`,
  `ComboBox`, `RadioGroup`, `TransferList`, or `CheckList` can now be
  fed straight from a `ListOf`, a `Queryable` result, or a
  `Json.Deserialize` array, alongside the array, Collection, and Range
  sources already supported.
- Requires ROneCOne 1.8.1 or later, a drop-in patch over 1.8.0 that
  fixes two Object-and-Variant dereference defects. Neither reached
  ReDim as written, but the same defect class did, in the readers
  above.

## 0.19.1 - 2026-07-28

- The spinner re-pins its frame after every rotation tick. A spinner
  was observed walking down and right during long multi-feed runs in
  Mission Control. Three live probes could not reproduce any frame
  movement at the object-model level: thousands of raw
  IncrementRotation calls, rotation under repaints, 90 and 125 percent
  zoom, interleaved shape writes with ScreenUpdating cycles, UIO
  protection, frame writes on a rotated shape, and the real demo
  under pumped load all held the frame exactly. The animation tick
  now snaps Left and Top back to the model whenever they deviate, so
  any walk, whatever drives it in an interactive session, corrects
  within a single frame, and the corrective write forces a repaint at
  the true rect. A new scenario proves a deviated frame is repaired
  in one frame.

## 0.19.0 - 2026-07-28

- Sprite downloads stopped blocking the pump. ReDex called
  `URLDownloadToFileA` synchronously inside the op that applied a
  species, so the first view of any sprite stalled a frame while the
  image came down. It now runs as its own async op on
  `HttpClient.DownloadFileAsync`, chasing the newest selection the same
  way the detail fetch does, and painting only if the dex still shows
  the species the image belongs to. The last `Declare` statement is
  gone from the demos.
- A species already seen paints from disk with no request at all, and
  the previous sprite stays up while a new one downloads, which reads
  better than flashing a placeholder. The placeholder rule itself is
  unchanged: a picture embedded in the workbook still outlives its
  source file.
- The demos stop hand-rolling what the runtime already provides. The
  ReDex detail cache is a `DictionaryOf(vbString, vbObject)` asked for
  membership directly, rather than a bare `Collection` probed by
  trapping a failed lookup. Temp paths come from `Path.Combine` and
  `Path.GetTempPath`, existence from `File.Exists`, and removal from
  `File.Delete`, in both ReDex and the Widget Gallery.

## 0.18.4 - 2026-07-28

- ReDex reads PokeAPI through ROneCOne 1.8.0's partial reads. The
  textual section stripper 0.18.3 shipped is gone, replaced by
  `Json.DeserializeOnly` with an allowlist of the seven paths the
  browse window actually draws. The unwanted move lists and per-game
  sprite variants are now stepped over inside the reader instead of
  cut out of the text beforehand, which is the same saving without a
  denylist that encoded one API's response shape. Measured for scale:
  responses run 271 KB to 665 KB and the dex draws under 2.5 KB of
  them.
- The detail cache holds parsed documents rather than payload strings,
  so a revisit costs neither a request nor a parse. Because an
  allowlist omits a path the response did not carry, the sprite reader
  checks for its members instead of assuming them.
- Requires ROneCOne 1.8.0 or later.

## 0.18.3 - 2026-07-27

- Theme swaps repaint everything. SetTheme now repaints a canvas that
  PrepareCanvas painted, so dark mode no longer leaves dark controls
  floating on the old light background, and the fills that were only
  styled on first apply (the progress track, the non-float input
  border, the anchored input cell) retint when the theme changes. A
  new core scenario locks the contract in.
- Progress bars skip their fill-part shape writes when nothing about
  the fill would change, keyed on value, color, visibility, and
  geometry. Tween loops and paced feeds stop paying for repaints of
  settled bars.
- ReDex switches fast. Detail payloads are slimmed before parsing
  (the unused move lists and per-game sprite variants carry most of a
  couple hundred KB, and the VBA JSON parse was the slowest link),
  then cached by request key, id, and name, so revisits apply
  instantly with no HTTP and rapid steps only parse what they show.
  Switching is latest-wins: a busy op records the newest selection
  and chases it when it settles instead of dropping the input, and
  F2/F3 advance a nav cursor on the keypress itself so held keys walk
  the dex. The stat tween runs at 30fps with faster convergence.

## 0.18.2 - 2026-07-26

- ReDex builds itself on open. The other demos already carried an
  Auto_Open; the new workbook shipped without one, so a fresh open
  showed a blank sheet until BuildPokeDex ran by hand. It now matches
  the family: open with macros enabled and the app draws itself.
  Automation opens skip Auto_Open by Excel's own rules, so the test
  harness keeps deciding when builds run.

## 0.18.1 - 2026-07-26

- ReDex: a fifth demo workbook and the full-framework showcase. A
  living Pokedex fed by PokeAPI over ROneCOne HTTP tasks riding the
  pump, parsed with ROneCOne JSON: a windowed 151-species combo with
  live filtering and pagers, sprite downloads into Image controls,
  base-stat bars that tween through a paced job, type badges restyled
  at runtime, F2/F3 hotkeys, a cancellable shiny hunt, a team window
  with the TransferList party builder and a house-rules CheckList, a
  trainer card with a Confirm-guarded reset, a custom Pokedex-red
  theme with a night-mode toggle flipping every window live, and
  toasts throughout. Offline it degrades honestly to an error toast
  and a retry on next show. The smoke test asserts structure,
  navigation, theming, and the fetch-op kick without depending on the
  network.

## 0.18.0 - 2026-07-26

- Hardening: the pump's silent error handling becomes observable and
  isolated. TickApp splits into guarded sections (ops, jobs, the
  component watch loop with toast expiry), so a fault in one duty can
  never silently skip the rest - the failure shape behind the
  long-lived compaction bug. Every silently trapped fault, including
  key-dispatch faults, now increments a factory counter exposed as
  `TickFaultCount` with `ResetTickFaults`; the toast slots scenario
  asserts zero across its dozens of forced ticks, so the suite proves
  the swallowed-error silence was earned.
- Performance: the focused field's selection poll (three COM reads)
  rides the 50 ms work cadence instead of every frame; TransferList
  and CheckList repaints are signature-skipped, so a touch that
  changes nothing visible costs no shape writes at all - both panels,
  headers, rows, buttons, and fonts previously rewrote on every touch.

## 0.17.3 - 2026-07-26

- Fix a press on floating chrome also engaging the slider underneath
  (reported: clicking the combo's pager row set the slider it
  overlapped). The slider's press watch is a geometric hit test with
  no notion of layering, so floating chrome now claims the points it
  covers: an open drop list claims its face, rows, and pagers, and a
  modal overlay claims everything it spans. The watch asks before
  engaging and skips claimed presses - which also closes the latent
  version of the same bug for sliders under modals.

## 0.17.2 - 2026-07-26

- The combo's inert overflow row is replaced by clickable pagers, on
  design feedback: the dead-end row gave mouse users no path to the
  overflow. Both list edges now page - a bottom row with a down arrow
  and the count below, and, once scrolled, a top row with an up arrow
  and the count above - keeping the list open, clamping at the ends,
  and swapping roles at the clamps. The same arrow vocabulary as the
  transfer panels. A Down after mouse paging starts the keyboard
  highlight at the window's first visible row instead of snapping back
  to the top, and the superseded indicator part is swept from adopted
  sheets.

## 0.17.1 - 2026-07-26

- Gallery long-list fixtures: the fruit combo grows to twenty items so
  the open list shows the eight-row window, the overflow indicator,
  the highlight walk, and live narrowing; the crew roster grows to
  fourteen so the available panel pages through its scroll arrows.

## 0.17.0 - 2026-07-26

- Long lists become navigable. The combo drop list windows to eight
  rows: Up and Down walk a keyboard highlight that scrolls the window,
  Enter takes the highlighted match, typing re-filters and rewinds, an
  inert indicator row counts the matches outside the window, and
  clicked rows map through the scroll offset. Transfer panels gain
  paging arrows on their right edge (shown only when the list outgrows
  the panel) that move the window a page at a time, with row selection
  mapped through each panel's offset; the header counts stay the
  truth-tellers for totals.

## 0.16.1 - 2026-07-26

- Esc on a combo clears instead of reverting: the first Esc empties the
  value, reopens the unfiltered list, and keeps focus for a fresh
  search; a second Esc on the empty combo blurs and commits the clear,
  so state never disagrees with the face. TextInput keeps the
  revert-on-Esc convention.

## 0.16.0 - 2026-07-26

- Full caret editing in float fields. The append-and-backspace model is
  retired: a caret index now lives anywhere in the buffer, Left and
  Right move it, Up and Down move it across hard lines with the column
  clamped to the target line, Home and End jump the line edges, Del
  deletes forward, Backspace deletes backward, and characters insert at
  the caret. The blinking bar renders at the caret's position, and the
  overflow viewport follows the caret in both directions - scrolling up
  shows a trailing ellipsis for the lines below. While a field is
  focused the arrow keys belong to editing, so they no longer move the
  cell selection and no longer trigger the selection-change commit;
  Home, End, and Del join the captured set and release on blur.
- Honest limits: vertical moves work in hard lines (a long wrapped line
  is one line to the caret), and apps that bind arrow HotKeys should
  re-arm them after field focus sessions if they mix the two.

## 0.15.2 - 2026-07-26

- Focused fields follow the caret when their text overflows. Shape text
  cannot scroll, so the focused view renders the tail window of the
  buffer instead: the last lines that fit a multi-line field, the
  rightmost characters that fit a single-line field or combo, always
  including the caret, with a leading ellipsis marking trimmed content.
  Window sizes are estimated conservatively from the font size, the
  buffer keeps the complete text, and commits are untouched. Unfocused
  fields still show their beginning.

## 0.15.1 - 2026-07-26

- Typing hot path slimmed, addressing keystroke lag and the busy-cursor
  blink. Measured first: replacing shape text preserves uniform
  character formatting, so the defensive font reapply on every text
  change is gone (kept only across empty-text transitions, the one case
  the measurement did not cover). The float field's fill and focus ring
  are diff-cached, and the combo option list skips its shape work when
  a keystroke leaves the visible list unchanged. Steady-state typing is
  now a single COM write per keystroke. The busy-cursor flip itself is
  Excel's macro indicator - every keystroke runs a macro by design -
  but with the light path it lasts a frame. A render debounce was
  considered and rejected: delaying the character echo trades a cursor
  blink for perceived input latency.
- The gallery notes field was two points too short for its second line;
  a line that cannot fully fit is not drawn at all. The field grew, and
  the MultiLine docs now carry the sizing rule of thumb.

## 0.15.0 - 2026-07-26

- Multi-line float TextInput: `.MultiLine` switches a field to textarea
  conventions - Enter inserts a newline and keeps focus, Tab and the
  newly bound Ctrl+Enter commit, clicking away commits, Esc reverts.
  Multi-line fields anchor their text to the top; the committed value
  carries its newlines into state and bindings. Single-line fields keep
  the Enter-commits convention unchanged. The gallery gains a notes
  field beside the combo.

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
