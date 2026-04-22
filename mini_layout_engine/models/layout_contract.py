from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping


SUPPORTED_FIELD_TYPES = {"string", "paragraph_list", "bullet_list", "image_ref"}


def _coerce_name_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value.strip()] if value.strip() else []
    if not isinstance(value, (list, tuple)):
        raise ValueError("Expected a list of field names.")
    names: list[str] = []
    for item in value:
        name = str(item or "").strip()
        if name:
            names.append(name)
    return names


@dataclass(frozen=True)
class LayoutArchetypeContract:
    layout_id: str
    description: str
    required_fields: list[str] = field(default_factory=list)
    optional_fields: list[str] = field(default_factory=list)
    field_types: dict[str, str] = field(default_factory=dict)
    constraints: dict[str, Any] = field(default_factory=dict)
    normalization_rules: dict[str, Any] = field(default_factory=dict)
    hook_key: str | None = None

    @property
    def all_fields(self) -> list[str]:
        seen: set[str] = set()
        ordered: list[str] = []
        for name in [*self.required_fields, *self.optional_fields]:
            if name in seen:
                continue
            seen.add(name)
            ordered.append(name)
        return ordered

    @classmethod
    def from_spec(cls, spec: Mapping[str, Any]) -> "LayoutArchetypeContract":
        if not isinstance(spec, Mapping):
            raise ValueError("Layout spec must be a JSON/YAML object.")

        layout_id = str(spec.get("layout_id", "")).strip()
        if not layout_id:
            raise ValueError("Layout spec is missing required field 'layout_id'.")

        description = str(spec.get("description", "")).strip()
        required_fields = _coerce_name_list(spec.get("required_fields", []))
        optional_fields = _coerce_name_list(spec.get("optional_fields", []))

        overlap = set(required_fields) & set(optional_fields)
        if overlap:
            overlap_list = ", ".join(sorted(overlap))
            raise ValueError(
                f"Layout '{layout_id}' has fields in both required/optional: {overlap_list}"
            )

        field_types_raw = spec.get("field_types", {})
        if field_types_raw is None:
            field_types_raw = {}
        if not isinstance(field_types_raw, Mapping):
            raise ValueError(
                f"Layout '{layout_id}' has invalid 'field_types'; expected object."
            )

        field_types = {str(k): str(v).strip() for k, v in field_types_raw.items()}
        for field_name in [*required_fields, *optional_fields]:
            if field_name not in field_types:
                field_types[field_name] = "string"

        constraints = spec.get("constraints", {})
        if constraints is None:
            constraints = {}
        if not isinstance(constraints, Mapping):
            raise ValueError(
                f"Layout '{layout_id}' has invalid 'constraints'; expected object."
            )

        normalization_rules = spec.get("normalization_rules", {})
        if normalization_rules is None:
            normalization_rules = {}
        if not isinstance(normalization_rules, Mapping):
            raise ValueError(
                f"Layout '{layout_id}' has invalid 'normalization_rules'; expected object."
            )

        hook_value = spec.get("hook_key")
        hook_key = str(hook_value).strip() if hook_value is not None else None
        if hook_key == "":
            hook_key = None

        return cls(
            layout_id=layout_id,
            description=description,
            required_fields=required_fields,
            optional_fields=optional_fields,
            field_types=field_types,
            constraints=dict(constraints),
            normalization_rules=dict(normalization_rules),
            hook_key=hook_key,
        )
