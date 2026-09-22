"""Stamps every release source with the release header.

Each VBA source a release ships (both runtime files and every demo module)
opens with the ReDim version and its release date, the repository, and the
MIT license text, as comments ahead of Option Explicit. The version comes
from REDIM_VERSION in ReDimUI.cls, the date from that version's CHANGELOG
entry, and the text from LICENSE.

Run it after bumping the version and adding the CHANGELOG entry. It replaces
whatever header a file carries, so running it twice changes nothing.
tests/python/test_source_guards.py fails until every source is stamped.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_sources import RELEASE_SOURCES, release_header  # noqa: E402

# The export preamble the header follows: the class VERSION/BEGIN/END block
# (its property lines are indented) and the Attribute lines.
PREAMBLE_PREFIXES = ("VERSION ", "BEGIN", "END", "Attribute ", "  ")


def stamp(path: Path, header: str) -> bool:
    raw = path.read_text(encoding="utf-8")
    newline = "\r\n" if "\r\n" in raw else "\n"
    lines = raw.split(newline)
    option_at = next(
        index for index, line in enumerate(lines)
        if line.startswith("Option Explicit")
    )
    keep = 0
    for index, line in enumerate(lines[:option_at]):
        if line.startswith(PREAMBLE_PREFIXES):
            keep = index + 1
    stamped = lines[:keep] + header.rstrip("\n").split("\n") + [""] + lines[option_at:]
    if stamped == lines:
        return False
    path.write_text(newline.join(stamped), encoding="utf-8", newline="")
    return True


def main() -> int:
    header = release_header()
    for path in RELEASE_SOURCES:
        changed = stamp(path, header)
        print(f"{path.name}: {'stamped' if changed else 'current'}")
    print(header.split("\n", 1)[0].lstrip("' "))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
