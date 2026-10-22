"""Dev lock gate: requirements-dev.txt keeps the packages only Windows needs.

The live suite runs on Windows, so the dev lock must install there with
--require-hashes. Dependabot recompiles it with pip-compile on Linux, which
resolves for Linux alone and drops every package only Windows needs; pip then
refuses the lock on Windows, since the packages that depend on them name them
without a pin. This gate fails such a lock and prints the command that
compiles it for every platform.
"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LOCK = ROOT / "requirements-dev.txt"
COMPILE = ("uv pip compile requirements-dev.in --universal --generate-hashes "
           "--python-version 3.12 -o requirements-dev.txt")

# Each package only Windows needs, and what pulls it in.
WINDOWS_PACKAGES = {
    "pywin32": "pyvbaharness, which drives Excel through COM",
    "colorama": "pytest",
    "win-unicode-console": "pcodedmp, which oletools pulls in",
}

PIN = re.compile(r"^([A-Za-z0-9][A-Za-z0-9._-]*)==\S+\s*;\s*(.*?)\s*\\?\s*$")


def missing_windows_packages(lock_text: str) -> list[str]:
    """The packages only Windows needs that the lock does not pin for Windows."""
    pinned = set()
    for line in lock_text.splitlines():
        match = PIN.match(line)
        if match and "sys_platform == 'win32'" in match.group(2):
            pinned.add(re.sub(r"[-_.]+", "-", match.group(1)).lower())
    return [name for name in WINDOWS_PACKAGES if name not in pinned]


def main() -> int:
    missing = missing_windows_packages(LOCK.read_text(encoding="utf-8"))
    if not missing:
        print(f"{LOCK.name} pins its Windows packages: {', '.join(WINDOWS_PACKAGES)}")
        return 0
    for name in missing:
        print(f"{LOCK.name} does not pin {name} for Windows, where "
              f"{WINDOWS_PACKAGES[name]} needs it")
    print("It was compiled for one platform, as Dependabot's pip-compile does. "
          f"Recompile it with: {COMPILE}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
