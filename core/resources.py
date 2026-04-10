"""PyInstaller-safe resource path helpers."""

from __future__ import annotations

import sys
from pathlib import Path


def resource_path(relative_path: str | Path) -> Path:
    """Resolve a resource path in development and frozen builds."""

    path = Path(relative_path)
    if path.is_absolute():
        return path
    return _resource_base() / path


def get_resource_path(relative_path: str | Path) -> Path:
    """Compatibility alias used by packaging docs."""

    return resource_path(relative_path)


def _resource_base() -> Path:
    frozen_base = getattr(sys, "_MEIPASS", None)
    if frozen_base:
        return Path(frozen_base)
    return Path(__file__).resolve().parent.parent


__all__ = ["resource_path", "get_resource_path"]
