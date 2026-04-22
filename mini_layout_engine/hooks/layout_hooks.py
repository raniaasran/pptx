from __future__ import annotations

from typing import Any


def run_layout_hook(
    *,
    hook_key: str | None,
    design_family_id: str,
    layout_id: str,
    normalized_content: dict[str, Any],
) -> dict[str, Any]:
    _ = normalized_content
    return {
        "status": "not_run",
        "hook_key": hook_key,
        "message": (
            "Layout customization hooks are intentionally disabled in V1 planning mode."
            f" design_family_id={design_family_id}, layout_id={layout_id}"
        ),
    }
