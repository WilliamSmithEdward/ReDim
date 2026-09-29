"""Fetch a checksum-pinned YARA Forge core bundle or propose its next release."""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PIN = ROOT / "tools/yara_forge_pin.json"
RULES = ROOT / "build/yara-forge/core.yar"
REPOSITORY = "YARAHQ/yara-forge"
ASSET = "yara-forge-rules-core.zip"
MEMBER = "packages/core/yara-rules-core.yar"
MAX_ARCHIVE = 20_000_000
MAX_RULES = 30_000_000


def download(url: str, token: str | None = None) -> bytes:
    headers = {"User-Agent": "ReDim-security-updater", "Accept": "application/octet-stream"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
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


def fetch() -> None:
    pin = json.loads(PIN.read_text(encoding="utf-8"))
    if pin["asset"] != ASSET or not re.fullmatch(r"[0-9a-f]{64}", pin["sha256"]):
        raise ValueError("invalid YARA Forge pin")
    data = download(release_asset(pin["release"]))
    digest = hashlib.sha256(data).hexdigest()
    if digest != pin["sha256"]:
        raise ValueError(f"YARA Forge SHA-256 mismatch: expected {pin['sha256']}, got {digest}")
    rules = unpack(data)
    RULES.parent.mkdir(parents=True, exist_ok=True)
    RULES.write_bytes(rules)
    print(f"Verified YARA Forge {pin['release']} ({digest}); {len(rules)} rule bytes")


def update() -> None:
    token = os.environ.get("GITHUB_TOKEN")
    headers = {"User-Agent": "ReDim-security-updater", "Accept": "application/vnd.github+json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    url = f"https://api.github.com/repos/{REPOSITORY}/releases/latest"
    with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=30) as response:
        release = json.load(response)
    tag = release["tag_name"]
    asset = next((item for item in release["assets"] if item["name"] == ASSET), None)
    if asset is None or asset["browser_download_url"] != release_asset(tag):
        raise ValueError("latest YARA Forge release has no expected core asset")
    current = json.loads(PIN.read_text(encoding="utf-8"))
    if tag == current["release"]:
        print(f"YARA Forge {tag} is already pinned")
        return
    data = download(asset["browser_download_url"])
    unpack(data)
    digest = hashlib.sha256(data).hexdigest()
    PIN.write_text(json.dumps({"release": tag, "asset": ASSET, "sha256": digest}, indent=2) + "\n",
                   encoding="utf-8")
    print(f"Proposed YARA Forge {tag} ({digest})")
    output = os.environ.get("GITHUB_OUTPUT")
    if output:
        with open(output, "a", encoding="utf-8") as stream:
            stream.write(f"release={tag}\nsha256={digest}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("fetch", "update"))
    args = parser.parse_args()
    fetch() if args.command == "fetch" else update()
