# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/ReDim/security/advisories/new).
Only the maintainer sees it. Include the release you used, the workbook or
module involved, and the smallest file or steps that show the problem,
with credentials and private data removed.

A confirmed vulnerability is fixed in a release on the GitHub releases
page, and the advisory is published with it,
crediting you unless you ask otherwise.

## Supported versions

Only the latest release on the GitHub releases page receives security
fixes. Older releases are not maintained separately; update when a fix
ships. ReDim is released from `main`.

## Scope

ReDim is VBA. Like any macro, it runs inside Excel with the rights of the
person who opens the workbook. A vulnerability here is ReDim's code, or
the demo workbooks, reaching something beyond what is described below.

`ReDimUI.cls` and `ReDimHost.bas`:

- call Windows API functions: `SetTimer` and `KillTimer` for the frame
  timer and `timeBeginPeriod` and `timeEndPeriod` for its resolution;
  `GetFocus`, `GetKeyState`, `GetClassNameW`, `PostMessageW`, and
  `MapVirtualKeyW` to route keys; gdi32 to measure text; user32 and
  kernel32 for the clipboard, the pointer, and settings such as the caret
  blink time; and `RegGetValueW` to read two values under
  `HKEY_CURRENT_USER` that hold the Windows dark mode and accent color;
- bind keys with `Application.OnKey` for an app's hot keys and shortcuts,
  and while a ReDim control has focus or a protected surface holds the
  typing keys; `Unmount`, `Shutdown`, and closing the workbook release
  them;
- call the handler procedures an app names, such as `OnClick` and
  `OnChange`, with `Application.Run`;
- read and write the clipboard for copy and paste in fields and tables;
- start no program, open no network connection, and write no file.

### The ROneCOne runtime

`ROneCOne.cls`, which ReDim needs and every demo workbook embeds, is a
general-purpose runtime. Its Process surface can start a program
(`cmd.exe` through `WScript.Shell`, or `CreateProcessW`), its HttpClient
makes HTTP requests (WinHttp), its File and Directory surfaces read and
write files, its Data surface uses ADODB, and its Hash surface uses
Windows CNG (`bcrypt.dll`). Each runs only when an app calls it. ReDim
itself uses ROneCOne's tasks, cancellation tokens, and JSON. Report a
vulnerability in the runtime itself to
[ROneCOne](https://github.com/WilliamSmithEdward/ROneCOne/security/advisories/new).

### The demo workbooks

The demo workbooks build themselves in `Auto_Open`. Two of them go
further:

- ReDex (`ReDim_ReDex.xlsm`) reads Pokemon data and sprite images from
  PokeAPI over HTTPS and caches the sprites in the temp folder as
  `redex_<id>.png`.
- Widget Gallery (`ReDim_Widget_Gallery.xlsm`) draws a logo on a chart
  and saves it to the temp folder as `rdm_gallery_logo.png`.

The other demos reach no network and write no file.

### Opening the demo workbooks

Releases no longer carry the demo workbooks. Build them from `demo/vba/`
with `python tools/build_workbooks.py`; a workbook built on your own
machine is not marked as downloaded. Office blocks macros in a workbook
downloaded from an earlier release: after checking its hash against that
release's security report, open its Properties in File Explorer, select
**Unblock**, and then open it and enable content.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily at 08:17 UTC.

Each scan covers `ReDimUI.cls`, `ReDimHost.bas`, the `ROneCOne.cls` the
build uses, and the demo workbooks, which `tools/build_workbooks.py`
builds fresh from source with pyOpenVBA for the run.

- **Code:** olevba and mraptor (MacroRaptor), both from oletools, and a
  p-code check, run by `tools/security_scan.py`. olevba and mraptor match
  keywords such as "run", "open", and "call" anywhere in the code,
  comments included, and every finding and mraptor match must be listed
  for its module with a reason. ReDim builds its workbooks from VBA source
  and Excel compiles the source when a workbook opens, so a workbook holds
  no compiled p-code. P-code that disagrees with its source is how VBA
  stomping hides code, so any p-code fails the scan, and olevba's own
  stomping check runs as well. The results go to the run's `macro-report`
  artifact, not to code scanning. CI also runs pyVBAanalysis.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** there is nothing to audit. ReDim's only dependency is
  ROneCOne, a VBA class checked out at a pinned commit (see Pinning and
  updates), and the Python tools the workflows run come from hash-locked
  files.
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan each source file and workbook as a file and then the
  VBA modules extracted from them. YARA-X also scans the extracted source
  with the repository's rules in `tools/vba_malware.yar`: broad rules for
  automatic execution, native API calls, process and network access, file
  writing, and dynamic invocation, and focused rules for encoded
  PowerShell, remote execution through Windows binaries, and Run key
  persistence. A YARA Forge download that does not match its pinned
  SHA-256, an unlisted match or detection, or a scanner error fails
  Malware scan. Signature updates can change the results, so the report
  records the engine and database version used.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Two of its checks do not fit this project. A single maintainer cannot
  have a second person approve every change. Fuzzing does not apply:
  ReDim is VBA, which runs only inside Office. Signed-Releases rises as
  releases carry the provenance bundle; it counts the last five.

A clean scan cannot establish that a workbook is safe.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[`tools/security_expected.json`](tools/security_expected.json). An entry
matches the tool, the module (or, for YARA Forge and ClamAV, the file or
module), and the finding: an olevba finding, an mraptor match, a YARA-X
rule identifier, or a ClamAV signature name. Entries are not tied to a
file's content. In the Security and Malware scan runs (`--strict`), an
entry that no longer matches fails the scan, so the list stays current.
zizmor keeps its exceptions in `.github/zizmor.yml` or inline beside the
line they excuse, each with its reason.

Current entries:

- olevba findings for 12 modules (`ThisWorkbook`, `Sheet1`, `ReDimUI`,
  `ReDimHost`, `ROneCOne` and the seven demo modules) and mraptor matches
  for 10, each naming one of 34 recorded reasons.
- mraptor calls every ReDim workbook SUSPICIOUS. It flags VBA that runs on
  its own (A) and also writes a file or memory (W) or runs code outside VBA
  (X). The demos build themselves in `Auto_Open`, the runtime declares
  Windows API functions, and ROneCOne carries its File and Process
  surfaces, so each workbook shows all three. `ReDimUI.cls` and
  `ROneCOne.cls` read "Macro OK": nothing in them runs on its own.
- YARA-X rule matches from the repository's broad rules for 10 modules.
  The focused rules have no entries.
- YARA Forge and ClamAV: there are none. The initial CI scan found no
  ClamAV signatures.
- zizmor: `.github/zizmor.yml` turns off `superfluous-actions`, which asks
  Publish to create the release with `gh release` in place of
  softprops/action-gh-release, until that change can be dry-run. There are
  no inline exceptions.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, the ClamAV image to a digest, Python tools
(oletools, YARA-X, pyOpenVBA, zizmor, pyVBAanalysis) to hash-locked lock
files in `.github/requirements/`, the local development tools to the
hash-locked `requirements-dev.txt`, and the YARA Forge rules to a release
and its SHA-256 in `.github/security/yara.json`. ClamAV's signatures change
too often to pin, so freshclam fetches and verifies them on every run.

The `ROneCOne.cls` that CI analyzes, Security and Malware scan build into
the demo workbooks, and Publish releases is checked out from the ROneCOne
repository at a full commit SHA (v1.10.3). No updater follows it; it moves
by hand in the workflows, and the comment at the pin in `ci.yml` gives the
commands.

Dependabot proposes updates to GitHub Actions, the Python lock files
(`.github/requirements/` and `requirements-dev.txt`) and the ClamAV image
once a version is a week old (the owner's own packages, such as pyOpenVBA
and pyVBAanalysis, at once), and at once for a security advisory. The
Update YARA rules workflow proposes new YARA pins each week. A minor or
patch update, and the YARA pull request, merges itself once CI, Security
and Malware scan pass; a third-party major version waits for review.

## Releases

Pushing a `vX.Y.Z` tag runs the Publish workflow
(`.github/workflows/publish.yml`). It refuses a tag that is not
`REDIM_VERSION` in `ReDimUI.cls`, runs `tools/stamp_release.py` and fails
if that changes any committed source, and takes the release files from the
tagged commit with the CRLF endings the VBE expects. It then runs
`tools/security_scan.py` over those files with ClamAV and the pinned YARA
Forge rules, judging findings by `tools/security_expected.json` at the
tag, and only when the scan passes signs the files' build provenance and
creates the GitHub release with:

- `ReDimUI.cls` and `ReDimHost.bas`: the runtime.
- `ROneCOne.cls`: the ROneCOne release the runtime is tested against,
  checked out at the commit pinned in the workflows.
- `ReDim-<version>.sigstore.json`: the signed build provenance.
- `vX.Y.Z-security-report.md`: the scan's report.

Releases no longer carry the demo workbooks; `tools/build_workbooks.py`
builds them from `demo/vba/` with pyOpenVBA. Releases up to 1.1.1 were
built locally, carry the workbooks, and have no provenance. Started by
hand, Publish is a dry run: it builds, scans and assembles the same files
and uploads them as the `release-preview` artifact instead.

### Release security reports

From 1.1.0, each release carries `vX.Y.Z-security-report.md`: the SHA-256
of every released file, mraptor's verdict on each file, each module's
olevba findings, mraptor matches, YARA-X (including pinned YARA Forge) and
ClamAV results, and the p-code check. The report names the workflow run
that wrote it. Reports up to 1.1.1 also cover the workbooks.

Workbooks up to 1.0.4 carry an empty `Module1` from the build template
whose p-code holds one comment line,
`TESTING ONLY DO NOT INCLUDE THIS IN FINAL OUTPUT`. It has no code and
never runs. 1.1.0 builds without it.

### Verifying a download

To check that a file was built by this repository's Publish workflow from
a tagged commit:

```bash
gh attestation verify ReDimUI.cls --repo WilliamSmithEdward/ReDim
```

Or compare a download's SHA-256 with the report, or with the digest GitHub
lists beside each asset:

```powershell
Get-FileHash .\ReDimUI.cls -Algorithm SHA256
```

To scan a file yourself:

```
pip install oletools==0.60.2 yara-x==1.20.0
olevba -a ReDimHost.bas
mraptor ReDimHost.bas
python tools/security_scan.py --clamav ReDimHost.bas  # requires clamscan and current signatures
```

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
