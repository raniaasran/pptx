from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping
from uuid import uuid4

from mini_layout_engine.hooks.layout_hooks import run_layout_hook
from mini_layout_engine.models.design_family import DesignFamilyContext
from mini_layout_engine.models.diagnostics import Diagnostics
from mini_layout_engine.models.slide_plan import (
    DeckPlanResult,
    PlaceholderStatus,
    SlidePlanResult,
)
from mini_layout_engine.models.slide_request import DeckPlanRequest, SlideInput
from mini_layout_engine.registry.family_registry import FamilyRegistry
from mini_layout_engine.registry.layout_registry import LayoutRegistry
from mini_layout_engine.validation.content_normalizer import normalize_content
from mini_layout_engine.validation.content_validator import validate_content


class PlanningEngine:
    def __init__(
        self,
        *,
        layout_registry: LayoutRegistry,
        family_registry: FamilyRegistry,
    ):
        self.layout_registry = layout_registry
        self.family_registry = family_registry

    @classmethod
    def from_default_specs(cls, specs_root: Path | None = None) -> "PlanningEngine":
        package_root = Path(__file__).resolve().parents[1]
        root = specs_root or package_root / "specs"
        layout_registry = LayoutRegistry.from_specs_dir(root / "layouts")
        family_registry = FamilyRegistry.from_specs_dir(root / "families")
        return cls(layout_registry=layout_registry, family_registry=family_registry)

    def plan_deck(self, request: DeckPlanRequest | Mapping[str, Any]) -> DeckPlanResult:
        deck_request = self._coerce_request(request)
        deck_diagnostics = Diagnostics()
        family_context = self._resolve_family_context(
            design_family_id=deck_request.design_family_id,
            diagnostics=deck_diagnostics,
        )

        theme_color_variant = self._resolve_variant(
            requested=deck_request.theme_color_variant,
            family_default=family_context.default_theme_color_variant,
        )
        font_variant = self._resolve_variant(
            requested=deck_request.font_variant,
            family_default=family_context.default_font_variant,
        )

        slide_results = [
            self._plan_slide(
                slide_input=slide_input,
                family_context=family_context,
                theme_color_variant=theme_color_variant,
                font_variant=font_variant,
            )
            for slide_input in deck_request.slides
        ]

        return DeckPlanResult(
            request_id=deck_request.request_id,
            design_family_id=family_context.family_id,
            template_file=family_context.template_file,
            theme_color_variant=theme_color_variant,
            font_variant=font_variant,
            theme_tokens=dict(family_context.theme_tokens),
            slides=slide_results,
            diagnostics=deck_diagnostics,
        )

    def _coerce_request(self, request: DeckPlanRequest | Mapping[str, Any]) -> DeckPlanRequest:
        if isinstance(request, DeckPlanRequest):
            return request
        if isinstance(request, Mapping):
            return DeckPlanRequest.from_dict(request)
        raise ValueError("plan_deck request must be a DeckPlanRequest or JSON object.")

    def _resolve_variant(self, *, requested: str | None, family_default: str | None) -> str:
        request_value = str(requested).strip() if requested is not None else ""
        if request_value:
            return request_value
        default_value = str(family_default).strip() if family_default is not None else ""
        if default_value:
            return default_value
        return "default"

    def _resolve_family_context(
        self,
        *,
        design_family_id: str,
        diagnostics: Diagnostics,
    ) -> DesignFamilyContext:
        family_context = self.family_registry.get(design_family_id)
        if family_context is not None:
            return family_context

        diagnostics.add_error(f"Unknown design_family_id '{design_family_id}'.")
        return DesignFamilyContext(
            family_id=str(design_family_id or "unknown"),
            template_file="unknown.pptx",
            default_theme_color_variant="default",
            default_font_variant="default",
            theme_tokens={},
            layout_mapping={},
        )

    def _plan_slide(
        self,
        *,
        slide_input: SlideInput,
        family_context: DesignFamilyContext,
        theme_color_variant: str,
        font_variant: str,
    ) -> SlidePlanResult:
        slide_diagnostics = Diagnostics()
        slide_instance_id = slide_input.slide_instance_id or f"slide_{uuid4().hex[:12]}"
        input_content = dict(slide_input.content)

        contract = self.layout_registry.get(slide_input.layout_id)
        if contract is None:
            slide_diagnostics.add_error(
                f"Unknown layout_id '{slide_input.layout_id}'. Registered layouts: "
                + ", ".join(self.layout_registry.all_layout_ids())
            )
            placeholder_status = PlaceholderStatus(
                required_fields=[],
                optional_fields=[],
                present_fields=sorted(input_content.keys()),
                missing_required_fields=[],
                unknown_fields=sorted(input_content.keys()),
                is_ready=False,
            )
            return SlidePlanResult(
                slide_instance_id=slide_instance_id,
                design_family_id=family_context.family_id,
                template_file=family_context.template_file,
                theme_color_variant=theme_color_variant,
                font_variant=font_variant,
                layout_id=slide_input.layout_id,
                mapped_layout={},
                readiness="not_ready",
                input_content=input_content,
                validated_content={},
                normalized_content={},
                placeholder_status=placeholder_status,
                adaptation_hook_status="not_run",
                adaptation_hook_key=None,
                diagnostics=slide_diagnostics,
            )

        validation_result = validate_content(contract, input_content)
        slide_diagnostics.extend(
            errors=validation_result.errors,
            warnings=validation_result.warnings,
        )

        mapped_layout = family_context.layout_mapping.get(contract.layout_id)
        if mapped_layout is None:
            slide_diagnostics.add_error(
                f"Layout '{contract.layout_id}' is not mapped in family "
                f"'{family_context.family_id}' layout_mapping."
            )
            mapped_layout = {}

        normalized_content = normalize_content(contract, validation_result.validated_content)
        hook_info = run_layout_hook(
            hook_key=contract.hook_key,
            design_family_id=family_context.family_id,
            layout_id=contract.layout_id,
            normalized_content=normalized_content,
        )

        placeholder_status = PlaceholderStatus(
            required_fields=list(contract.required_fields),
            optional_fields=list(contract.optional_fields),
            present_fields=sorted(validation_result.validated_content.keys()),
            missing_required_fields=list(validation_result.missing_required_fields),
            unknown_fields=list(validation_result.unknown_fields),
            is_ready=not slide_diagnostics.has_errors
            and not validation_result.missing_required_fields,
        )

        readiness = "ready" if placeholder_status.is_ready else "not_ready"

        return SlidePlanResult(
            slide_instance_id=slide_instance_id,
            design_family_id=family_context.family_id,
            template_file=family_context.template_file,
            theme_color_variant=theme_color_variant,
            font_variant=font_variant,
            layout_id=contract.layout_id,
            mapped_layout=dict(mapped_layout),
            readiness=readiness,
            input_content=input_content,
            validated_content=dict(validation_result.validated_content),
            normalized_content=normalized_content,
            placeholder_status=placeholder_status,
            adaptation_hook_status=str(hook_info.get("status", "not_run")),
            adaptation_hook_key=contract.hook_key,
            diagnostics=slide_diagnostics,
        )
