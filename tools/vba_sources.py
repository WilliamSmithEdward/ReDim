"""Shared helpers for reading VBA sources and locating the ROneCOne runtime.

The line-ending and class-preamble handling follows the ROneCOne build pipeline so
injected modules round-trip byte for byte through pyOpenVBA.
"""

from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
TESTS_VBA = ROOT / "tests" / "vba"
DEMO_VBA = ROOT / "demo" / "vba"
OUTPUT = ROOT / "tests" / "output"


def ronecone_class_path() -> Path:
    """Locate ROneCOne.cls: env override first, then the sibling checkout."""
    override = os.environ.get("REDIM_RONECONE_PATH")
    if override:
        path = Path(override)
        if not path.is_file():
            raise FileNotFoundError(f"REDIM_RONECONE_PATH does not exist: {path}")
        return path
    sibling = ROOT.parent / "ROneCOne" / "src" / "ROneCOne.cls"
    if sibling.is_file():
        return sibling
    raise FileNotFoundError(
        "ROneCOne.cls not found. Clone ROneCOne next to ReDim or set REDIM_RONECONE_PATH."
    )


def read_vba(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def prepare_class_source(path: Path) -> str:
    """Remove the VBE export-only VERSION/BEGIN preamble for binary injection."""
    source = read_vba(path)
    lines = source.split("\r\n")
    if not lines or lines[0].strip().upper() != "VERSION 1.0 CLASS":
        return source
    for index, line in enumerate(lines[1:], start=1):
        if line.strip().upper() == "END":
            return "\r\n".join(lines[index + 1 :])
    raise ValueError(f"Class export preamble has no END marker: {path}")
