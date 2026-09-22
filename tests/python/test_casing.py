"""Identifier casing guard: importing ReDim must not recase the names in a
host project's modules.

VBA keeps one spelling per identifier across a whole project, and a
declaration anywhere sets it: a parameter named `value` in ReDimUI.cls
turns every `.Value` in the host's modules into `.value`, which lands as
noise in every exported diff. The static tests hold the runtime and the
demos to three rules over their code tokens: one spelling per name, the
type-library spelling for any name the default references define, and
ROneCOne's spelling for any name it exposes as a member. The live test
checks the outcome directly: modules exported through the VBE come back
token for token as they went in.

ROneCOne stays out of the round-trip workbook. It declares names such as
`value` and `text` in lowercase itself, which is its own defect to fix;
with it present, the export would measure ROneCOne rather than ReDim.
"""

from __future__ import annotations

import functools
import re
import shutil
from collections import defaultdict
from pathlib import Path

import pytest
from pyopenvba import ExcelFile, VBAModuleKind
from pyvbaharness import ExcelSession

from vba_sources import (
    DEMO_VBA,
    OUTPUT,
    SRC,
    prepare_class_source,
    read_vba,
    ronecone_class_path,
)

RUNTIME = [SRC / "ReDimUI.cls", SRC / "ReDimHost.bas"]
DEMOS = sorted(DEMO_VBA.glob("*.bas"))

# The references every new Excel VBA project carries, in reference order.
REFERENCES = (
    "{000204EF-0000-0000-C000-000000000046}",  # VBA
    "{00020813-0000-0000-C000-000000000046}",  # Excel
    "{00020430-0000-0000-C000-000000000046}",  # stdole
    "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}",  # Office
)

# Keywords keep their own spelling and never enter the name table, so a
# keyword that shares letters with a library member (GoTo, Goto) is fine.
KEYWORDS = {word.lower() for word in """
    AddressOf And Any As Attribute Base Binary Boolean ByRef Byte ByVal Call
    Case Compare Const Currency Database Date Decimal Declare Dim Do Double
    Each Else ElseIf Empty End Enum Eqv Erase Error Event Exit Explicit False
    For Friend Function Get Global GoSub GoTo If Imp Implements In Integer Is
    Let Lib Like Long LongLong LongPtr Loop LSet Me Mod Module New Next Not
    Nothing Null Object On Option Optional Or ParamArray Preserve Print
    Private Property PtrSafe Public RaiseEvent ReDim Rem Resume Return RSet
    Select Set Single Static Step Stop String Sub Text Then To True Type
    TypeOf Until Variant Wend While With WithEvents Xor
""".split()}

WORD = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
NUMBER = re.compile(
    r"&[Hh][0-9A-Fa-f]+&?|&[Oo][0-7]+&?|"
    r"(?<![A-Za-z0-9_])\d+(?:\.\d+)?(?:[eE][+-]?\d+)?[#!@&%^]?")
MEMBER = re.compile(
    r"^\s*(?:Public|Friend)\s+(?:Static\s+)?"
    r"(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.MULTILINE)

# Host code as a ReDim user writes it: Excel members, named arguments, an
# event-style Target parameter, and ReDim's own members, with nothing
# declared that shares their names. Every token must survive the export.
HOST_PROBE = """\
Attribute VB_Name = "HostProbe"
Option Explicit

Public Sub BuildHostProbe()
    Dim ws As Worksheet
    Dim ui As ReDimUI

    Set ws = ActiveSheet
    ws.Range("A1").Value = 1
    ws.Range("A2").Formula = "=A1*2"
    Debug.Print ws.Range("A1").Text, ws.Cells(1, 1).Value, ws.Name
    Debug.Print ThisWorkbook.Names.Count, ThisWorkbook.FullName, ThisWorkbook.Title
    ws.Range("A1").Interior.Color = vbRed
    ws.Range("A1").Font.Name = "Arial"
    ws.Range("A1").Font.Size = 10
    ws.Cells.Item(RowIndex:=1, ColumnIndex:=2).Value = 3
    ws.Hyperlinks.Add Anchor:=ws.Range("B1"), Address:=""
    Debug.Print Application.Caller, Application.Hwnd, ActiveWindow.Caption
    Debug.Print Err.Source, Err.Description, Err.Number, ws.Index
    Debug.Print ws.Shapes.Item(1).Visible, ws.Rows(1).RowHeight
    Set ui = ReDimUI.Mount(ws, "probe")
    Debug.Print ui.AppId, ui.Sheet.Name, ReDimUI.App("probe").AppId
    ui.Button("go").Text("Go").Visible(True).Enabled True
    Debug.Print ui.Component("go").ComponentId, ui.Component("go").OwnerApp.AppId
    Debug.Print ReDimUI.SenderId, ReDimUI.FocusedComponentId, ui.Theme.FontName
End Sub

Private Sub HandleChange(ByVal Target As Range, Cancel As Boolean)
    Debug.Print Target.Address, Target.Row, Target.Column, Cancel
End Sub
"""


def code_tokens(text: str) -> list[str]:
    """Identifier tokens after Option Explicit: comments, strings, and
    numeric literals are blanked first, so only code participates."""
    lines = text.replace("\r\n", "\n").split("\n")
    start = next(
        (i for i, line in enumerate(lines) if line.startswith("Option Explicit")),
        0,
    )
    found = []
    for line in lines[start:]:
        kept = []
        in_string = False
        for ch in line:
            if ch == '"':
                in_string = not in_string
                kept.append(" ")
            elif ch == "'" and not in_string:
                break
            else:
                kept.append(" " if in_string else ch)
        found.extend(WORD.findall(NUMBER.sub(" ", "".join(kept))))
    return found


def _registered_typelib(guid: str):
    import pythoncom
    import winreg

    def subkeys(path: str) -> list[str]:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, path) as key:
            return [winreg.EnumKey(key, i)
                    for i in range(winreg.QueryInfoKey(key)[0])]

    version = max(subkeys(rf"TypeLib\{guid}"),
                  key=lambda v: tuple(int(part, 16) for part in v.split(".")))
    lcid = next(int(name, 16) for name in subkeys(rf"TypeLib\{guid}\{version}")
                if re.fullmatch(r"[0-9A-Fa-f]+", name))
    major, minor = (int(part, 16) for part in version.split("."))
    return pythoncom.LoadRegTypeLib(guid, major, minor, lcid)


@functools.lru_cache(maxsize=1)
def typelib_spellings() -> dict[str, str]:
    """The spelling the default references give each name: a type, member,
    or constant spelling beats a parameter spelling, and references rank
    in project order."""
    best: dict[str, tuple[tuple[int, int], str]] = {}

    def offer(name: str, rank: tuple[int, int]) -> None:
        key = name.lower()
        if key not in best or rank < best[key][0]:
            best[key] = (rank, name)

    for order, guid in enumerate(REFERENCES):
        library = _registered_typelib(guid)
        for index in range(library.GetTypeInfoCount()):
            info = library.GetTypeInfo(index)
            offer(library.GetDocumentation(index)[0], (0, order))
            attr = info.GetTypeAttr()
            for f in range(attr.cFuncs):
                names = info.GetNames(info.GetFuncDesc(f).memid)
                for position, name in enumerate(names):
                    offer(name, (1 if position else 0, order))
            for v in range(attr.cVars):
                for name in info.GetNames(info.GetVarDesc(v).memid):
                    offer(name, (0, order))
    return {key: spelling for key, (_, spelling) in best.items()}


def ronecone_members() -> dict[str, str]:
    text = ronecone_class_path().read_text(encoding="utf-8")
    return {name.lower(): name for name in MEMBER.findall(text)}


def casing_problems(paths: list[Path]) -> list[str]:
    try:
        canon = typelib_spellings()
    except Exception as exc:  # noqa: BLE001 - any load failure means no oracle
        pytest.skip(f"default type libraries unavailable: {exc}")
    members = ronecone_members()
    spellings: dict[str, dict[str, list[str]]] = defaultdict(lambda: defaultdict(list))
    for path in paths:
        for token in code_tokens(path.read_text(encoding="utf-8")):
            spellings[token.lower()][token].append(path.name)
    problems = []
    for lower, forms in sorted(spellings.items()):
        names = sorted(forms)
        if len(names) > 1:
            problems.append(f"{lower}: spelled {names}")
        if lower in KEYWORDS:
            continue
        wanted = canon.get(lower) or members.get(lower)
        if wanted is not None:
            for name in names:
                if name != wanted:
                    where = sorted(set(forms[name]))
                    problems.append(f"{name} in {where}: {wanted} is canonical")
    return problems


def test_runtime_casing():
    problems = casing_problems(RUNTIME)
    assert not problems, (
        "ReDim's runtime would recase these names in a host project:\n"
        + "\n".join(problems)
    )


def test_demo_casing():
    problems = casing_problems(RUNTIME + DEMOS)
    assert not problems, (
        "the demos, shipped beside the runtime, would recase these names:\n"
        + "\n".join(problems)
    )


def test_vbe_export_round_trips():
    """The user's own scenario: a project holding ReDim, every demo, and
    host code exports through the VBE with every token spelled as
    written."""
    modules: dict[str, tuple[str, VBAModuleKind]] = {
        "ReDimUI": (prepare_class_source(SRC / "ReDimUI.cls"), VBAModuleKind.other),
        "ReDimHost": (read_vba(SRC / "ReDimHost.bas"), VBAModuleKind.standard),
    }
    for path in DEMOS:
        modules[path.stem] = (read_vba(path), VBAModuleKind.standard)
    modules["HostProbe"] = (
        HOST_PROBE.replace("\n", "\r\n"), VBAModuleKind.standard)
    workbook = OUTPUT / "casing_roundtrip.xlsm"
    export_dir = OUTPUT / "casing_roundtrip_export"
    workbook.parent.mkdir(parents=True, exist_ok=True)
    if workbook.exists():
        workbook.unlink()
    shutil.rmtree(export_dir, ignore_errors=True)
    with ExcelFile.create_new(workbook) as book:
        project = book.vba_project()
        for name, (source, kind) in modules.items():
            project.add_module(name, source, kind=kind)
        book.save()
    with ExcelSession() as excel:
        excel.open_workbook(str(workbook))
        excel.export_modules(export_dir)
    changed = []
    for name, (source, _) in modules.items():
        exported = next(export_dir.glob(f"{name}.*")).read_text(
            encoding="utf-8", errors="replace")
        before, after = code_tokens(source), code_tokens(exported)
        assert len(before) == len(after), f"{name}: token count changed"
        pairs = sorted({(a, b) for a, b in zip(before, after) if a != b})
        changed.extend(f"{name}: {a} -> {b}" for a, b in pairs)
    assert not changed, "the VBE recased these tokens:\n" + "\n".join(changed)
