# Security

## Reporting a vulnerability

Report it privately: on the repository's Security tab choose **Report a vulnerability**, or go
straight to <https://github.com/WilliamSmithEdward/ReDim/security/advisories/new>. The report
stays between you and the maintainer until an advisory is published. Please do not describe a
vulnerability in a public issue.

## Supported versions

Fixes go into the next release. ReDim is released from `main` and earlier versions get no
patches, so use the [latest release](https://github.com/WilliamSmithEdward/ReDim/releases).

## What the code does

ReDim is VBA. Like any macro, it runs inside Excel with the rights of the person who opens the
workbook.

`ReDimUI.cls` and `ReDimHost.bas`:

- call Windows API functions: `SetTimer` and `KillTimer` for the frame timer and
  `timeBeginPeriod` and `timeEndPeriod` for its resolution; `GetFocus`, `GetKeyState`,
  `GetClassNameW`, `PostMessageW`, and `MapVirtualKeyW` to route keys; gdi32 to measure text;
  user32 and kernel32 for the clipboard, the pointer, and settings such as the caret blink time;
  and `RegGetValueW` to read two values under `HKEY_CURRENT_USER` that hold the Windows dark mode
  and accent color;
- bind keys with `Application.OnKey` for an app's hot keys and shortcuts, and while a ReDim
  control has focus or a protected surface holds the typing keys; `Unmount`, `Shutdown`, and
  closing the workbook release them;
- call the handler procedures an app names, such as `OnClick` and `OnChange`, with
  `Application.Run`;
- read and write the clipboard for copy and paste in fields and tables;
- start no program, open no network connection, and write no file.

`ROneCOne.cls`, which ReDim needs and every demo workbook embeds, is a general-purpose runtime.
Its Process surface can start a program (`cmd.exe` through `WScript.Shell`, or
`CreateProcessW`), its HttpClient makes HTTP requests (WinHttp), its File and Directory surfaces
read and write files, its Data surface uses ADODB, and its Hash surface uses Windows CNG
(`bcrypt.dll`). Each runs only when an app calls it. ReDim itself uses ROneCOne's tasks,
cancellation tokens, and JSON.

The demo workbooks build themselves in `Auto_Open`. Two of them go further:

- ReDex (`ReDim_ReDex.xlsm`) reads Pokemon data and sprite images from PokeAPI over HTTPS and
  caches the sprites in the temp folder as `redex_<id>.png`.
- Widget Gallery (`ReDim_Widget_Gallery.xlsm`) draws a logo on a chart and saves it to the temp
  folder as `rdm_gallery_logo.png`.

The other demos reach no network and write no file.

## What scanners report

Macro scanners flag much of the above, and ordinary words too: olevba and mraptor match
keywords such as "run", "open", and "call" anywhere in the code, comments included.
[`tools/security_expected.json`](tools/security_expected.json) lists every olevba finding and
mraptor match for each module with the reason for it. On each push, the
[Security workflow](.github/workflows/security.yml) runs
[`tools/security_scan.py`](tools/security_scan.py) over the runtime files and freshly built demo
workbooks, and fails on any finding the list does not explain, and on an expected one that no
longer occurs.

mraptor (MacroRaptor) calls every ReDim workbook SUSPICIOUS. It flags VBA that runs on its own
(A) and also writes a file or memory (W) or runs code outside VBA (X). The demos build
themselves in `Auto_Open`, the runtime declares Windows API functions, and ROneCOne carries its
File and Process surfaces, so each workbook shows all three. `ReDimUI.cls` alone reads "Macro
OK": nothing in it runs on its own.

The scan also checks each workbook for p-code. ReDim builds its workbooks from VBA source with
pyOpenVBA, and Excel compiles the source when a workbook opens, so a workbook holds no compiled
p-code. P-code that disagrees with its source is how VBA stomping hides code, and any p-code
fails the scan. Workbooks up to 1.0.4 carry an empty `Module1` from the build template whose
p-code holds one comment line, `TESTING ONLY DO NOT INCLUDE THIS IN FINAL OUTPUT`. It has no
code and never runs. 1.1.0 builds without it.

## Release security reports

From 1.1.0, each release carries `vX.Y.Z-security-report.md`. The
[release security workflow](.github/workflows/release-security.yml) writes it from the
release's published files: the SHA-256 of every workbook and source file, mraptor's verdict on
each file, each module's olevba findings and mraptor matches with their reasons, and the p-code
check. The report names the workflow run that wrote it.

To check a download, compare its SHA-256 with the report, or with the digest GitHub lists
beside each asset:

```powershell
Get-FileHash .\ReDim_Snake.xlsm -Algorithm SHA256
```

To scan a file yourself:

```
pip install oletools
olevba -a ReDim_Snake.xlsm
mraptor ReDim_Snake.xlsm
```

## Opening the demo workbooks

Office blocks macros in files downloaded from the internet. After checking a workbook's hash,
open its Properties in File Explorer, select **Unblock**, and then open it and enable content.
