from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping


@dataclass(frozen=True)
class SlideInput:
    layout_id: str
    content: dict[str, Any]
    slide_instance_id: str | None = None

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "SlideInput":
        if not isinstance(value, Mapping):
            raise ValueError("Each slide request must be an object.")

        layout_id = str(value.get("layout_id", "")).strip()
        if not layout_id:
            raise ValueError("Slide request is missing required field 'layout_id'.")

        raw_content = value.get("content", {})
        if raw_content is None:
            raw_content = {}
        if not isinstance(raw_content, Mapping):
            raise ValueError(
                f"Slide '{layout_id}' has invalid 'content'; expected object."
            )

        slide_instance_id = value.get("slide_instance_id")
        instance_value = (
            str(slide_instance_id).strip() if slide_instance_id is not None else None
        )
        if instance_value == "":
            instance_value = None

        return cls(
            layout_id=layout_id,
            content={str(k): v for k, v in dict(raw_content).items()},
            slide_instance_id=instance_value,
        )


@dataclass(frozen=True)
class DeckPlanRequest:
    design_family_id: str
    slides: list[SlideInput]
    theme_color_variant: str | None = None
    font_variant: str | None = None
    request_id: str | None = None

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "DeckPlanRequest":
        if not isinstance(value, Mapping):
            raise ValueError("Planning request must be a JSON object.")

        design_family_id = str(value.get("design_family_id", "")).strip()
        if not design_family_id:
            raise ValueError("Planning request is missing 'design_family_id'.")

        raw_slides = value.get("slides", [])
        if not isinstance(raw_slides, list):
            raise ValueError("Planning request field 'slides' must be a list.")

        slides = [SlideInput.from_dict(item) for item in raw_slides]

        theme_color_variant_raw = value.get("theme_color_variant")
        font_variant_raw = value.get("font_variant")
        request_id_raw = value.get("request_id")

        theme_color_variant = (
            str(theme_color_variant_raw).strip() if theme_color_variant_raw is not None else None
        )
        font_variant = str(font_variant_raw).strip() if font_variant_raw is not None else None
        request_id = str(request_id_raw).strip() if request_id_raw is not None else None

        if theme_color_variant == "":
            theme_color_variant = None
        if font_variant == "":
            font_variant = None
        if request_id == "":
            request_id = None

        return cls(
            design_family_id=design_family_id,
            slides=slides,
            theme_color_variant=theme_color_variant,
            font_variant=font_variant,
            request_id=request_id,
        )
