from __future__ import annotations

import json
from pathlib import Path
from typing import Any


SUPPORTED_SPEC_SUFFIXES = (".json",)


def iter_spec_paths(directory: Path) -> list[Path]:
    if not directory.exists():
        return []
    paths = [
        path
        for path in directory.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_SPEC_SUFFIXES
    ]
    return sorted(paths, key=lambda path: path.name.lower())


def load_spec_file(path: Path) -> dict[str, Any]:
    raw_text = path.read_text(encoding="utf-8").strip()
    if not raw_text:
        raise ValueError(f"Spec file is empty: {path}")

    try:
        value = json.loads(raw_text)
    except json.JSONDecodeError:
        raise ValueError(f"Spec file '{path.name}' is not valid JSON.") from None

    if not isinstance(value, dict):
        raise ValueError(f"Spec file must contain an object at root: {path}")
    return value
