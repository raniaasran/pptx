from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .diagnostics import Diagnostics


@dataclass
class PlaceholderStatus:
    required_fields: list[str] = field(default_factory=list)
    optional_fields: list[str] = field(default_factory=list)
    present_fields: list[str] = field(default_factory=list)
    missing_required_fields: list[str] = field(default_factory=list)
    unknown_fields: list[str] = field(default_factory=list)
    is_ready: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "required_fields": list(self.required_fields),
            "optional_fields": list(self.optional_fields),
            "present_fields": list(self.present_fields),
            "missing_required_fields": list(self.missing_required_fields),
            "unknown_fields": list(self.unknown_fields),
            "is_ready": bool(self.is_ready),
        }


@dataclass
class SlidePlanResult:
    slide_instance_id: str
    design_family_id: str
    template_file: str
    theme_color_variant: str
    font_variant: str
    layout_id: str
    readiness: str
    mapped_layout: dict[str, Any] = field(default_factory=dict)
    input_content: dict[str, Any] = field(default_factory=dict)
    validated_content: dict[str, Any] = field(default_factory=dict)
    normalized_content: dict[str, Any] = field(default_factory=dict)
    placeholder_status: PlaceholderStatus = field(default_factory=PlaceholderStatus)
    adaptation_hook_status: str = "not_run"
    adaptation_hook_key: str | None = None
    diagnostics: Diagnostics = field(default_factory=Diagnostics)

    def to_dict(self) -> dict[str, Any]:
        return {
            "slide_instance_id": self.slide_instance_id,
            "design_family_id": self.design_family_id,
            "template_file": self.template_file,
            "theme_color_variant": self.theme_color_variant,
            "font_variant": self.font_variant,
            "layout_id": self.layout_id,
            "mapped_layout": dict(self.mapped_layout),
            "readiness": self.readiness,
            "input_content": dict(self.input_content),
            "validated_content": dict(self.validated_content),
            "normalized_content": dict(self.normalized_content),
            "placeholder_status": self.placeholder_status.to_dict(),
            "adaptation_hook_status": self.adaptation_hook_status,
            "adaptation_hook_key": self.adaptation_hook_key,
            "diagnostics": self.diagnostics.to_dict(),
        }


@dataclass
class DeckPlanResult:
    design_family_id: str
    template_file: str
    theme_color_variant: str
    font_variant: str
    theme_tokens: dict[str, Any] = field(default_factory=dict)
    slides: list[SlidePlanResult] = field(default_factory=list)
    diagnostics: Diagnostics = field(default_factory=Diagnostics)
    request_id: str | None = None

    @property
    def status(self) -> str:
        if self.diagnostics.has_errors:
            return "error"
        if any(slide.diagnostics.has_errors for slide in self.slides):
            return "partial"
        return "ok"

    def to_dict(self) -> dict[str, Any]:
        return {
            "request_id": self.request_id,
            "status": self.status,
            "design_family_id": self.design_family_id,
            "template_file": self.template_file,
            "theme_color_variant": self.theme_color_variant,
            "font_variant": self.font_variant,
            "theme_tokens": dict(self.theme_tokens),
            "slides": [slide.to_dict() for slide in self.slides],
            "diagnostics": self.diagnostics.to_dict(),
            "summary": {
                "slides_total": len(self.slides),
                "slides_ready": sum(1 for slide in self.slides if slide.readiness == "ready"),
                "slides_not_ready": sum(
                    1 for slide in self.slides if slide.readiness != "ready"
                ),
            },
        }
