"""Helpers to check GitHub releases and launch platform installers."""

from __future__ import annotations

import json
import os
import re
import stat
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Iterable

GITHUB_REPO = "CSRuger/Klausurmaster"
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
WINDOWS_SUFFIXES: tuple[str, ...] = (".exe", ".msi")
MAC_SUFFIXES: tuple[str, ...] = (".dmg", ".pkg", ".zip")
LINUX_SUFFIXES: tuple[str, ...] = (".AppImage", ".tar.gz", ".sh")


class UpdateError(RuntimeError):
    """Raised when update steps fail."""


def _normalize_version(value: str) -> tuple[int, ...]:
    cleaned = value.strip().lstrip("vV")
    if not cleaned:
        return ()
    parts = re.split(r"[._-]", cleaned)
    normalized = []
    for part in parts:
        if part.isdigit():
            normalized.append(int(part))
    return tuple(normalized)


def fetch_latest_release(timeout: int = 10) -> dict[str, Any]:
    try:
        with urllib.request.urlopen(API_URL, timeout=timeout) as response:
            data = response.read()
    except urllib.error.URLError as exc:
        raise UpdateError(f"GitHub Anfrage fehlgeschlagen: {exc}") from exc
    try:
        payload = json.loads(data.decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise UpdateError("Antwort von GitHub konnte nicht gelesen werden") from exc
    return payload


def is_newer_version(current: str, latest: str) -> bool:
    return _normalize_version(latest) > _normalize_version(current)


def _pick_suffixes() -> Iterable[str]:
    if sys.platform.startswith("win"):
        return WINDOWS_SUFFIXES
    if sys.platform == "darwin":
        return MAC_SUFFIXES
    return LINUX_SUFFIXES


def select_best_asset(release_payload: dict[str, Any]) -> dict[str, Any] | None:
    assets = release_payload.get("assets") or []
    suffixes = _pick_suffixes()
    for suffix in suffixes:
        for asset in assets:
            name = asset.get("name") or ""
            if name.endswith(suffix):
                return asset
    return None


def download_asset(asset: dict[str, Any], directory: Path | None = None) -> Path:
    url = asset.get("browser_download_url")
    name = asset.get("name")
    if not url or not name:
        raise UpdateError("Asset-Informationen unvollständig")
    target_dir = directory or Path(tempfile.gettempdir())
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = target_dir / name
    try:
        with urllib.request.urlopen(url) as response, target_path.open("wb") as handle:
            handle.write(response.read())
    except urllib.error.URLError as exc:
        raise UpdateError(f"Download fehlgeschlagen: {exc}") from exc
    return target_path


def launch_installer(installer_path: Path) -> None:
    if sys.platform.startswith("win"):
        subprocess.Popen([str(installer_path)], shell=False)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(installer_path)], shell=False)
    else:
        current_mode = os.stat(installer_path).st_mode
        installer_path.chmod(current_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
        subprocess.Popen([str(installer_path)], shell=False)