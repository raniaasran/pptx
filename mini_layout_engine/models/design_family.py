from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping


@dataclass(frozen=True)
class DesignFamilyContext:
    family_id: str
    template_file: str
    default_theme_color_variant: str = "default"
    default_font_variant: str = "default"
    theme_tokens: dict[str, Any] = field(default_factory=dict)
    layout_mapping: dict[str, dict[str, Any]] = field(default_factory=dict)

    @classmethod
    def from_spec(cls, spec: Mapping[str, Any]) -> "DesignFamilyContext":
        if not isinstance(spec, Mapping):
            raise ValueError("Family spec must be a JSON object.")

        family_id = str(spec.get("family_id", "")).strip()
        if not family_id:
            raise ValueError("Family spec is missing required field 'family_id'.")

        template_file = str(spec.get("template_file", "")).strip()
        if not template_file:
            raise ValueError(f"Family '{family_id}' is missing required field 'template_file'.")

        theme_tokens = spec.get("theme_tokens", {})
        if theme_tokens is None:
            theme_tokens = {}
        if not isinstance(theme_tokens, Mapping):
            raise ValueError(
                f"Family '{family_id}' has invalid 'theme_tokens'; expected object."
            )

        default_theme_color_variant = str(
            spec.get("default_theme_color_variant", "default")
        ).strip() or "default"
        default_font_variant = str(spec.get("default_font_variant", "default")).strip() or "default"

        layout_mapping_raw = spec.get("layout_mapping", {})
        if layout_mapping_raw is None:
            layout_mapping_raw = {}
        if not isinstance(layout_mapping_raw, Mapping):
            raise ValueError(
                f"Family '{family_id}' has invalid 'layout_mapping'; expected object."
            )

        layout_mapping: dict[str, dict[str, Any]] = {}
        for layout_id, mapping in layout_mapping_raw.items():
            key = str(layout_id).strip()
            if not key:
                continue
            if not isinstance(mapping, Mapping):
                raise ValueError(
                    f"Family '{family_id}' mapping for layout '{key}' must be an object."
                )
            layout_mapping[key] = {str(k): v for k, v in dict(mapping).items()}

        return cls(
            family_id=family_id,
            template_file=template_file,
            default_theme_color_variant=default_theme_color_variant,
            default_font_variant=default_font_variant,
            theme_tokens=dict(theme_tokens),
            layout_mapping=layout_mapping,
        )
