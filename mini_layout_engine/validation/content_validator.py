from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping

from mini_layout_engine.models.layout_contract import LayoutArchetypeContract


@dataclass
class ContentValidationResult:
    validated_content: dict[str, Any] = field(default_factory=dict)
    missing_required_fields: list[str] = field(default_factory=list)
    unknown_fields: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.errors


def _is_empty_value(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, (list, tuple)):
        if len(value) == 0:
            return True
        return all(_is_empty_value(item) for item in value)
    return False


def _matches_field_type(value: Any, field_type: str) -> bool:
    normalized_type = str(field_type or "").strip().lower()
    if normalized_type in {"string", "image_ref"}:
        return isinstance(value, str)
    if normalized_type in {"paragraph_list", "bullet_list"}:
        if isinstance(value, str):
            return True
        if isinstance(value, (list, tuple)):
            return all(item is None or isinstance(item, (str, int, float, bool)) for item in value)
        return False
    return True


def validate_content(
    contract: LayoutArchetypeContract,
    content: Mapping[str, Any] | None,
) -> ContentValidationResult:
    result = ContentValidationResult()

    if content is None:
        content_map: dict[str, Any] = {}
    elif isinstance(content, Mapping):
        content_map = {str(k): v for k, v in dict(content).items()}
    else:
        result.errors.append("Slide content must be an object.")
        return result

    allowed_fields = set(contract.all_fields)
    incoming_fields = set(content_map.keys())

    unknown = sorted(incoming_fields - allowed_fields)
    if unknown:
        result.unknown_fields = unknown
        result.warnings.append(
            f"Unknown fields ignored for layout '{contract.layout_id}': {', '.join(unknown)}"
        )

    for field_name in contract.required_fields:
        if field_name not in content_map or _is_empty_value(content_map[field_name]):
            result.missing_required_fields.append(field_name)

    if result.missing_required_fields:
        result.errors.append(
            f"Missing required fields for layout '{contract.layout_id}': "
            + ", ".join(result.missing_required_fields)
        )

    for field_name in contract.all_fields:
        if field_name not in content_map:
            continue
        value = content_map[field_name]
        result.validated_content[field_name] = value
        field_type = contract.field_types.get(field_name, "string")
        if not _matches_field_type(value, field_type):
            result.errors.append(
                f"Field '{field_name}' has invalid type for layout '{contract.layout_id}'. "
                f"Expected '{field_type}'."
            )

    return result
