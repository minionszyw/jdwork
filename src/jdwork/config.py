"""Shared configuration path helpers for the jdwork commands."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def default_config_path(command: str, cwd: Path | None = None) -> Path:
    """Return the project-local default configuration for a CLI command."""
    base = (cwd or Path.cwd()).resolve()
    return base / "config" / f"{command}.json"


def resolve_config_path(value: str | None, command: str, cwd: Path | None = None) -> Path:
    """Resolve an explicit or project-local configuration path."""
    path = Path(value) if value else default_config_path(command, cwd)
    if not path.is_absolute():
        path = (cwd or Path.cwd()) / path
    return path.resolve()


def source_path(root: Path, item: dict[str, Any] | str, key: str = "file") -> Path:
    value = item[key] if isinstance(item, dict) else item
    path = Path(value)
    return path if path.is_absolute() else root / path


def read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as handle:
        config = json.load(handle)
    if not isinstance(config, dict):
        raise ValueError(f"配置文件必须包含 JSON 对象: {path}")
    return config
