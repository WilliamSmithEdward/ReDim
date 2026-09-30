"""Fetch the checksum-pinned YARA Forge core bundle the scans run.

The pin is the `yara_forge` entry of .github/security/yara.json, which the
standard YARA updater (.github/security/yara_update.py) moves.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PIN = ROOT / ".github/security/yara.json"
RULES = ROOT / "build/yara-forge/core.yar"
REPOSITORY = "YARAHQ/yara-forge"
ASSET = "yara-forge-rules-core.zip"
MEMBER = "packages/core/yara-rules-core.yar"
MAX_ARCHIVE = 20_000_000
MAX_RULES = 30_000_000


def download(url: str) -> bytes:
    headers = {"User-Agent": "ReDim-security-scan", "Accept": "application/octet-stream"}
    with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=90) as response:
        data = response.read(MAX_ARCHIVE + 1)
    if len(data) > MAX_ARCHIVE:
        raise ValueError("YARA Forge archive exceeds size limit")
    return data


def release_asset(tag: str) -> str:
    if not re.fullmatch(r"\d{8}", tag):
        raise ValueError(f"unexpected YARA Forge release tag: {tag!r}")
    return f"https://github.com/{REPOSITORY}/releases/download/{tag}/{ASSET}"


def unpack(data: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        members = [entry for entry in archive.infolist() if not entry.is_dir()]
        if len(members) != 1 or members[0].filename != MEMBER:
            raise ValueError("unexpected YARA Forge core archive contents")
        if members[0].file_size > MAX_RULES:
            raise ValueError("YARA Forge rules exceed size limit")
        return archive.read(MEMBER)


def read_pin() -> dict:
    pin = json.loads(PIN.read_text(encoding="utf-8"))["yara_forge"]
    if (pin["asset"] != ASSET or pin["url"] != release_asset(pin["release"])
            or not re.fullmatch(r"[0-9a-f]{64}", pin["sha256"])):
        raise ValueError("invalid YARA Forge pin")
    return pin


def fetch() -> None:
    pin = read_pin()
    data = download(pin["url"])
    digest = hashlib.sha256(data).hexdigest()
    if digest != pin["sha256"]:
        raise ValueError(f"YARA Forge SHA-256 mismatch: expected {pin['sha256']}, got {digest}")
    rules = unpack(data)
    RULES.parent.mkdir(parents=True, exist_ok=True)
    RULES.write_bytes(rules)
    print(f"Verified YARA Forge {pin['release']} ({digest}); {len(rules)} rule bytes")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("fetch",))
    parser.parse_args()
    fetch()
