"""Security scan of ReDim's VBA with olevba, plus a p-code check.

CI runs it over the sources and freshly built demo workbooks on every push;
the release workflow runs it over a release's assets and attaches the report
it writes.

olevba (oletools) scans each module for auto-run entry points, keywords it
associates with malware, and indicators such as URLs and file names. Every
finding must be listed for its module in security_expected.json with the
reason ReDim has it, or the scan fails. A hex or base64 string passes when
it is 16 characters or fewer: a number or a short word that decodes by
chance. olevba scans every decoded string for its keywords as well.

A workbook's VBA project must hold no p-code lines. ReDim builds its
workbooks from source alone and Excel compiles them when they open, so
p-code in a workbook is code its source does not show, which is what VBA
stomping hides. olevba's own stomping check runs as well.

Usage: security_scan.py [--strict] [--report FILE] [--title TEXT] FILE [FILE ...]
Exit 0 when nothing unexpected turned up, 1 otherwise; with --strict, an
expected finding that matched nothing fails the scan too.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

# oletools is imported where it is used, so the parts that need only the
# standard library load, and test, without it.

EXPECTED_PATH = Path(__file__).resolve().parent / "security_expected.json"
SHORT_ENCODED = re.compile(r"[A-Za-z0-9+/=]{1,16}")
ENCODED_KINDS = ("Hex String", "Base64 String")
PCODE_MODULE = re.compile(r"^(?:VBA/)?(\S+) - \d+ bytes$")
WORKBOOK_SUFFIXES = (".xlsm", ".xlsb", ".xlam")


@dataclass
class Module:
    name: str
    code: str
    digest: str
    files: list[str] = field(default_factory=list)
    # (type, keyword, olevba's note, reason or None when unexpected)
    verdicts: list[tuple[str, str, str, str | None]] = field(default_factory=list)


@dataclass
class ScannedFile:
    path: Path
    digest: str
    size: int
    modules: list[str] = field(default_factory=list)
    pcode: dict | None = None


def sha256_of(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def plain(text: str) -> str:
    """Text any console prints: a decoded string can hold anything."""
    return text.encode("ascii", "backslashreplace").decode("ascii")


def module_code(code: str) -> str:
    """A module's code as a workbook stores it: LF line ends, and without
    the VERSION ... END preamble a class module's exported file carries, so
    ReDimUI.cls and the ReDimUI inside a workbook compare as one module."""
    lines = code.replace("\r\n", "\n").split("\n")
    if lines and lines[0].strip().upper() == "VERSION 1.0 CLASS":
        for index, line in enumerate(lines[1:], start=1):
            if line.strip().upper() == "END":
                lines = lines[index + 1:]
                break
    return "\n".join(lines)


def read_modules(path: Path) -> list[tuple[str, str]]:
    """Each module's name and code, as olevba extracts them."""
    from oletools import olevba

    parser = olevba.VBA_Parser(str(path))
    try:
        if not parser.detect_vba_macros():
            return []
        return [(Path(vba_filename).stem, module_code(code))
                for _, _, vba_filename, code in parser.extract_macros()]
    finally:
        parser.close()


def judge_module(module: Module, expected: dict) -> None:
    from oletools import olevba

    allowed = {key.casefold(): reason
               for key, reason in expected["modules"].get(module.name, {}).items()}
    reasons = expected["reasons"]
    # As olevba reports by default: a decoded string is listed when it reads
    # as text. Every decoded string is also scanned for keywords, and a hit
    # there is a finding of its own.
    scanner = olevba.VBA_Scanner(module.code)
    for kind, keyword, note in scanner.scan(include_decoded_strings=False):
        if kind in ENCODED_KINDS:
            # The note holds the encoded text, the keyword its decoding.
            reason = reasons["encoded"] if SHORT_ENCODED.fullmatch(note) else None
            module.verdicts.append((kind, note, f"reads as {keyword!r}", reason))
            continue
        reason_id = allowed.get(f"{kind}: {keyword}".casefold())
        module.verdicts.append((kind, keyword, note, reasons.get(reason_id)))


def pcode_lines(pcodedmp_output: str) -> dict[str, list[str]]:
    """The p-code instructions pcodedmp disassembled, by module. A module
    stored as source alone has none; pcodedmp reports an error for it."""
    found: dict[str, list[str]] = {}
    current = None
    in_modules = False
    for line in pcodedmp_output.splitlines():
        if line.strip() == "Module streams:":
            in_modules = True
            continue
        if not in_modules:
            continue
        header = PCODE_MODULE.match(line.strip())
        if header:
            current = header.group(1)
        elif current and line.startswith("\t") and line.strip():
            found.setdefault(current, []).append(line.strip())
    return found


def check_pcode(path: Path) -> dict:
    """Disassembles a workbook's VBA project and runs olevba's stomping
    check, which works only on a project file on disk."""
    from oletools import olevba

    with zipfile.ZipFile(path) as package:
        projects = [name for name in package.namelist()
                    if name.lower().endswith("vbaproject.bin")]
        if not projects:
            return {"error": "no VBA project in the package"}
        with tempfile.TemporaryDirectory() as work:
            target = Path(work) / "vbaProject.bin"
            target.write_bytes(package.read(projects[0]))
            parser = olevba.VBA_Parser(str(target))
            try:
                stomped = bool(parser.detect_vba_stomping())
                output = parser.pcodedmp_output or ""
            finally:
                parser.close()
    if "Module streams:" not in output:
        return {"error": "pcodedmp could not read the project's modules"}
    return {"stomped": stomped, "lines": pcode_lines(output)}


def tool_versions() -> dict[str, str]:
    from oletools import olevba
    from pcodedmp import pcodedmp

    return {"olevba": olevba.__version__, "pcodedmp": pcodedmp.__VERSION__}


def write_report(target: Path, title: str, files: list[ScannedFile],
                 modules: list[Module], problems: list[str], stale: list[str]) -> None:
    versions = tool_versions()
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    out: list[str] = [f"# {title}", "",
                      f"Scanned {stamp} with olevba {versions['olevba']} (oletools) "
                      f"and pcodedmp {versions['pcodedmp']}."]
    run_id = os.environ.get("GITHUB_RUN_ID")
    if run_id:
        server = os.environ.get("GITHUB_SERVER_URL", "https://github.com")
        repository = os.environ.get("GITHUB_REPOSITORY", "")
        out.append(f"Produced by {server}/{repository}/actions/runs/{run_id}.")
    out.append("")
    if problems:
        out.append(f"**{len(problems)} unexpected finding(s).** They are listed at the end.")
    else:
        out.append("**Nothing unexpected.** Every finding below is one ReDim expects, "
                   "with the reason it is there.")
    out += ["", "## Files", "",
            "| File | SHA-256 | Bytes | Modules |", "| --- | --- | ---: | --- |"]
    for scanned in files:
        names = ", ".join(scanned.modules) or "none"
        out.append(f"| {scanned.path.name} | `{scanned.digest}` | {scanned.size:,} | {names} |")
    workbooks = [scanned for scanned in files if scanned.pcode is not None]
    if workbooks:
        out += ["", "## Workbook p-code", "",
                "ReDim builds its workbooks from source alone, and Excel compiles the "
                "source when a workbook opens. P-code in a workbook would be code the "
                "source does not show, which is how VBA stomping hides it.", ""]
        for scanned in workbooks:
            pcode = scanned.pcode
            if "error" in pcode:
                out.append(f"- {scanned.path.name}: the check could not run: {pcode['error']}.")
                continue
            held = ", ".join(f"{name} ({len(lines)} lines)"
                             for name, lines in pcode["lines"].items()) or "none"
            stomping = "detected" if pcode["stomped"] else "not detected"
            out.append(f"- {scanned.path.name}: p-code lines: {held}; "
                       f"olevba's stomping check: {stomping}.")
    out += ["", "## olevba findings by module", "",
            "olevba matches its keywords as whole words anywhere in the code, comments "
            "included, so it also flags ordinary words. Each finding is grouped under "
            "the reason ReDim has it.", ""]
    for module in modules:
        out += [f"### {module.name}", "",
                f"{len(module.code.splitlines()):,} lines, SHA-256 `{module.digest}`, "
                f"in {', '.join(module.files)}.", ""]
        if not module.verdicts:
            out += ["No findings.", ""]
            continue
        groups: dict[str, list[str]] = {}
        for kind, keyword, note, reason in module.verdicts:
            groups.setdefault(reason or "**Unexpected.**", []).append(
                f"- {kind} `{keyword}`: {note}")
        for reason, lines in groups.items():
            out += [reason, "", *lines, ""]
    out += ["## Unexpected", ""]
    out += [f"- {problem}" for problem in problems] or ["None."]
    if stale:
        out += ["", "## Expected but not found", "",
                "These entries in tools/security_expected.json matched nothing:", ""]
        out += [f"- {entry}" for entry in stale]
    out += ["", "## Reproduce", "",
            "From a ReDim checkout, with the files to scan:", "",
            "```",
            f"pip install oletools=={versions['olevba']}",
            "olevba -a <file>",
            "python tools/security_scan.py <file> ...",
            "```", ""]
    target.write_text("\n".join(out), encoding="utf-8")


def main() -> int:
    options = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    options.add_argument("files", nargs="+", type=Path)
    options.add_argument("--strict", action="store_true",
                         help="also fail on expected findings that matched nothing")
    options.add_argument("--report", type=Path, help="write the Markdown report here")
    options.add_argument("--title", default="ReDim security scan")
    args = options.parse_args()

    expected = json.loads(EXPECTED_PATH.read_text(encoding="utf-8"))
    problems: list[str] = []
    files: list[ScannedFile] = []
    modules: dict[tuple[str, str], Module] = {}
    for path in args.files:
        data = path.read_bytes()
        scanned = ScannedFile(path, sha256_of(data), len(data))
        files.append(scanned)
        try:
            extracted = read_modules(path)
        except Exception as error:  # olevba raises many kinds on a bad file
            problems.append(f"{path.name}: olevba could not read it: {error}")
            continue
        if not extracted:
            problems.append(f"{path.name}: olevba found no VBA in it")
        for name, code in extracted:
            scanned.modules.append(name)
            if not code.strip():
                continue
            key = (name, sha256_of(code.encode("utf-8")))
            module = modules.setdefault(key, Module(name, code, key[1]))
            module.files.append(path.name)
        if path.suffix.lower() in WORKBOOK_SUFFIXES:
            scanned.pcode = check_pcode(path)
            if "error" in scanned.pcode:
                problems.append(f"{path.name}: the p-code check could not run: "
                                f"{scanned.pcode['error']}")
            else:
                for name, lines in scanned.pcode["lines"].items():
                    problems.append(f"{path.name}: module {name} holds p-code: {lines[0]}")
                if scanned.pcode["stomped"]:
                    problems.append(f"{path.name}: olevba detected VBA stomping")

    ordered = sorted(modules.values(), key=lambda module: (module.name.casefold(), module.digest))
    for module in ordered:
        judge_module(module, expected)
        for kind, keyword, note, reason in module.verdicts:
            if reason is None:
                problems.append(f"{module.name}: olevba {kind} {keyword!r}: {note}")

    seen = {module.name for module in ordered}
    found = {(module.name, f"{kind}: {keyword}".casefold())
             for module in ordered for kind, keyword, _, _ in module.verdicts}
    stale = [f"{name}: {key}" for name, entries in expected["modules"].items()
             if name in seen for key in entries if (name, key.casefold()) not in found]

    for scanned in files:
        print(f"{scanned.path.name}  {scanned.digest}  {', '.join(scanned.modules)}")
    for module in ordered:
        print(f"  {module.name}: {len(module.verdicts)} olevba findings")
    for entry in stale:
        print(plain(f"expected but not found: {entry}"))
    for problem in problems:
        print(plain(f"UNEXPECTED {problem}"))
    print(f"scanned {len(files)} files, {len(ordered)} modules, "
          f"unexpected: {len(problems)}")
    if args.report:
        write_report(args.report, args.title, files, ordered, problems, stale)
    return 1 if problems or (args.strict and stale) else 0


if __name__ == "__main__":
    raise SystemExit(main())
