"""Utilities for locating resources and persistent data paths across platforms."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

APP_NAME = "Klausurmaster"
APP_VERSION = "2.0.0"
CONFIG_FILENAME = "config.json"
DEFAULT_SAVE_FILENAME = "Tabellenspeicher_neu.json"
ASSETS_DIRNAME = "assets"


def _project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _assets_root() -> Path:
    return _project_root() / ASSETS_DIRNAME


def get_user_data_dir() -> Path:
    """Return the per-user data directory, honoring overrides and platform defaults."""
    override = os.environ.get("KLAUSURMASTER_DATA_DIR")
    if override:
        return Path(override).expanduser()

    home = Path.home()
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("APPDATA", home))
        return base / APP_NAME
    if sys.platform == "darwin":
        return home / "Library" / "Application Support" / APP_NAME
    return home / ".local" / "share" / APP_NAME


def get_user_config_path() -> Path:
    return get_user_data_dir() / CONFIG_FILENAME


def _load_config() -> dict[str, Any]:
    config_path = get_user_config_path()
    if config_path.exists():
        try:
            with config_path.open("r", encoding="utf-8") as handle:
                return json.load(handle)
        except (json.JSONDecodeError, OSError):
            pass

    template_path = _assets_root() / CONFIG_FILENAME
    if template_path.exists():
        try:
            with template_path.open("r", encoding="utf-8") as handle:
                return json.load(handle)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _persist_config(config: dict[str, Any]) -> None:
    data_dir = get_user_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    config_path = data_dir / CONFIG_FILENAME
    with config_path.open("w", encoding="utf-8") as handle:
        json.dump(config, handle, indent=2, ensure_ascii=False)


def _resolve_save_file(raw_value: str | None) -> Path:
    if raw_value:
        candidate = Path(raw_value).expanduser()
        if not candidate.is_absolute():
            candidate = get_user_data_dir() / candidate
    else:
        candidate = get_user_data_dir() / DEFAULT_SAVE_FILENAME
    candidate.parent.mkdir(parents=True, exist_ok=True)
    return candidate


def load_save_file_path() -> Path:
    """Return the save-file path, creating config defaults if necessary."""
    env_override = os.environ.get("KLAUSURMASTER_SAVE_FILE")
    config = _load_config()
    target = _resolve_save_file(env_override or config.get("save_file"))
    if env_override is None:
        config["save_file"] = str(target)
        _persist_config(config)
    return target


def persist_save_file_path(new_path: str | Path) -> str:
    """Persist a new save-file path and return it as a string."""
    path_obj = Path(new_path).expanduser()
    if not path_obj.is_absolute():
        path_obj = get_user_data_dir() / path_obj
    path_obj.parent.mkdir(parents=True, exist_ok=True)
    config = _load_config()
    config["save_file"] = str(path_obj)
    _persist_config(config)
    return str(path_obj)


def resource_path(relative_name: str) -> Path:
    """Return a path to a bundled resource (works for PyInstaller and source runs)."""
    base = Path(getattr(sys, "_MEIPASS", _project_root()))
    candidate = base / relative_name
    if candidate.exists():
        return candidate

    alt = base / ASSETS_DIRNAME / relative_name
    if alt.exists():
        return alt

    project_candidate = _project_root() / relative_name
    if project_candidate.exists():
        return project_candidate

    return _assets_root() / relative_name


def project_root() -> Path:
    return _project_root()
