"""Pure-source structural guards: invariants the VBA runtime must hold by
construction, checked by reading the source without launching Excel.

These lock in defect-class fixes at the shape level, not just the repaired
line, the same way ROneCOne guards its own repaired classes.
"""

from __future__ import annotations

import re
from pathlib import Path

from vba_sources import RELEASE_SOURCES, release_header

REDIMUI = Path(__file__).resolve().parents[2] / "src" / "ReDimUI.cls"
GUIDES = [
    Path(__file__).resolve().parents[2] / "docs" / name for name in ("api.md", "async.md")
]

PUBLIC_MEMBER = re.compile(
    r"^Public\s+(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)", re.M
)
# A fenced example or an inline span; the fence comes first so its
# backticks never pair up as spans.
CODE_SPAN = re.compile(r"```.*?```|`[^`\n]*`", re.DOTALL)

# A bare IsArray call, not IsArrayValue: "IsArray" followed by optional
# whitespace and an open paren. "IsArrayValue(" has "Value" after IsArray,
# so it never matches, and prose mentions of IsArray lack the paren.
ISARRAY_CALL = re.compile(r"\bIsArray\s*\(")
GUARD_FUNCTION = re.compile(
    r"Private Function IsArrayValue.*?End Function", re.DOTALL
)


def test_isarray_appears_once_inside_the_guard():
    text = REDIMUI.read_text(encoding="utf-8")
    calls = ISARRAY_CALL.findall(text)
    assert len(calls) == 1, (
        f"IsArray must appear exactly once in the runtime, found {len(calls)}. "
        "An object reaching a bare IsArray dereferences its default member and "
        "raises; route every array test through the IsArrayValue guard."
    )
    guard = GUARD_FUNCTION.search(text)
    assert guard is not None, "the IsArrayValue guard function must exist"
    assert len(ISARRAY_CALL.findall(guard.group(0))) == 1, (
        "the one IsArray call must live inside the IsArrayValue guard, "
        "which short-circuits on IsObject before testing"
    )


def test_every_public_member_is_documented():
    """A Public member of ReDimUI is API; the guides name each one in code,
    inline or in an example, so a new member cannot ship undocumented."""
    members = set(PUBLIC_MEMBER.findall(REDIMUI.read_text(encoding="utf-8")))
    documented = set()
    for guide in GUIDES:
        for span in CODE_SPAN.findall(guide.read_text(encoding="utf-8")):
            documented.update(re.findall(r"[A-Za-z_]\w*", span))
    missing = sorted(members - documented)
    assert not missing, (
        "Public members the guides never name in a code span: " + ", ".join(missing)
    )


CONTROL_OR_DECLARATION = re.compile(
    r"^(Set |Call |If |ElseIf |Else|End |For |Next|Do|Loop|While|Wend|Select |Case |"
    r"With |Dim |Private |Public |Friend |Const |ReDim |Exit |On Error|GoTo|Resume|"
    r"Static |Attribute |Option |Debug\.|Err\.|Erase |Stop|Declare )"
)
STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')


def _statements(path: Path):
    """Each logical line with the number it starts on: continuations joined,
    with no space before a continued member access."""
    buffer = ""
    start = 0
    lines = path.read_text(encoding="utf-8").replace("\r\n", "\n").split("\n")
    for number, line in enumerate(lines, 1):
        code = line.rstrip()
        if buffer:
            code = code.lstrip()
            if not code.startswith("."):
                code = " " + code
        else:
            start = number
        if code.endswith(" _"):
            buffer += code[:-2]
            continue
        yield start, (buffer + code).strip()
        buffer = ""


def _ends_in_parenthesized_call(statement: str) -> bool:
    """True for obj.Method(a, b) as a statement: the last group holds a comma
    and everything before it is a bare member chain, not an argument list."""
    bare = STRING_LITERAL.sub('""', statement)
    if not bare.endswith(")"):
        return False
    depth = 0
    for index in range(len(bare) - 1, -1, -1):
        if bare[index] == ")":
            depth += 1
        elif bare[index] == "(":
            depth -= 1
            if depth == 0:
                break
    else:
        return False
    head, inner = bare[:index], bare[index + 1:-1]
    if not re.search(r"[.\w]$", head) or re.search(r"(?<![<>=])=(?!=)", head.split("(")[0]):
        return False
    level = 0
    comma = False
    for char in inner:
        level += (char == "(") - (char == ")")
        comma = comma or (char == "," and level == 0)
    level = 0
    for char in head:
        level += (char == "(") - (char == ")")
        if char == " " and level == 0:
            return False
    return comma


def test_no_statement_ends_in_a_parenthesized_call():
    """VBA rejects a statement whose last call wraps several arguments in
    parentheses, ui.Card("c").AtRect(0, 0, 9, 9) for one, and only a compile
    in Excel reports it. The api.md statement-form chaining rule, checked
    across every VBA source."""
    root = REDIMUI.parents[1]
    sources = [REDIMUI, root / "src" / "ReDimHost.bas"]
    sources += sorted((root / "tests" / "vba").glob("*.bas"))
    sources += sorted((root / "demo" / "vba").glob("*.bas"))
    offenders = [
        f"{path.name}:{number}: {statement}"
        for path in sources
        for number, statement in _statements(path)
        if statement and not statement.startswith("'")
        and not CONTROL_OR_DECLARATION.match(statement)
        and _ends_in_parenthesized_call(statement)
    ]
    assert not offenders, (
        "Write the last call of a statement without parentheses, "
        "obj.Method a, b:\n" + "\n".join(offenders)
    )


def test_ui_text_table_follows_its_indexes():
    """UI_TEXT_TABLE lists the UI texts in the order of their UI_ index
    constants, one entry each: UI_NO_ROWS_MATCH is entry 2, NoRowsMatch.
    A table out of step would give every text after the slip the wrong
    words."""
    text = REDIMUI.read_text(encoding="utf-8").replace("\r\n", "\n")
    numbers = {
        int(number): name
        for name, number in re.findall(r"Private Const (UI_\w+) As Long = (\d+)", text)
        if name != "UI_TEXT_COUNT"
    }
    count = int(re.search(r"Private Const UI_TEXT_COUNT As Long = (\d+)", text).group(1))
    body = re.search(r"Private Const UI_TEXT_TABLE As String = _\n((?:.*\n)*?)(?!\s+\")",
                     text).group(1)
    table = "".join(re.findall(r'"((?:[^"]|"")*)"', body))
    names = [entry.split("=", 1)[0] for entry in table.split("|")]
    expected = [numbers[index] for index in range(1, count + 1)]
    derived = ["UI_" + re.sub(r"(?<=[a-z])(?=[A-Z])", "_", name).upper() for name in names]
    assert derived == expected, "UI_TEXT_TABLE is out of step with the UI_ constants"


def test_release_sources_carry_license_and_version():
    """Every source a release ships opens with the version and release date,
    the repository, and the MIT license, built from REDIM_VERSION, the
    CHANGELOG, and LICENSE; tools/stamp_release.py writes it."""
    header = release_header()
    for path in RELEASE_SOURCES:
        text = path.read_text(encoding="utf-8").replace("\r\n", "\n")
        preamble = text.split("\nOption Explicit", 1)[0]
        assert header in preamble, (
            f"{path.name} lacks the current release header; run "
            "python tools/stamp_release.py:\n" + header
        )
