from __future__ import annotations

from typing import Any, Mapping


SLIDE_KEY_PREFIX = "slide_"
TEMPLATE_KEY_FIELD = "template_key"
SLIDE_METADATA_FIELDS = {TEMPLATE_KEY_FIELD}


def _is_slide_key(value: str) -> bool:
    normalized = str(value or "").strip().lower()
    if not normalized.startswith(SLIDE_KEY_PREFIX):
        return False
    suffix = normalized[len(SLIDE_KEY_PREFIX) :]
    return suffix.isdigit()


def _template_placeholders(template_spec: Any) -> Mapping[str, Any]:
    if isinstance(template_spec, Mapping):
        placeholders = template_spec.get("placeholders")
        if isinstance(placeholders, Mapping):
            return placeholders
        return template_spec
    return {}


def _template_signature(template_spec: Any) -> tuple[str, ...]:
    placeholders = _template_placeholders(template_spec)
    return tuple(sorted(str(key) for key in placeholders.keys()))


def _is_picture_field(field_name: str) -> bool:
    return "picture" in str(field_name or "").lower()


def _collect_slide_field_values(slide_payload: Mapping[str, Any]) -> tuple[list[str], list[str]]:
    text_values: list[str] = []
    picture_values: list[str] = []
    for key, value in slide_payload.items():
        if str(key) in SLIDE_METADATA_FIELDS:
            continue
        cleaned_value = str(value if value is not None else "").strip()
        if not cleaned_value:
            continue
        if _is_picture_field(str(key)):
            picture_values.append(cleaned_value)
        else:
            text_values.append(cleaned_value)
    return text_values, picture_values


def _project_slide_payload_to_template(
    slide_payload: Mapping[str, Any],
    *,
    template_key: str,
    available_templates: Mapping[str, Any],
) -> dict[str, Any]:
    target_placeholders = _template_placeholders(available_templates.get(template_key))
    projected: dict[str, Any] = {TEMPLATE_KEY_FIELD: template_key}

    source_text_values, source_picture_values = _collect_slide_field_values(slide_payload)
    consumed_text_values: set[int] = set()
    consumed_picture_values: set[int] = set()

    for placeholder_key in target_placeholders.keys():
        existing_value = slide_payload.get(placeholder_key)
        cleaned_existing = str(existing_value if existing_value is not None else "").strip()
        if cleaned_existing:
            projected[str(placeholder_key)] = cleaned_existing
            if _is_picture_field(str(placeholder_key)):
                for index, candidate in enumerate(source_picture_values):
                    if index in consumed_picture_values:
                        continue
                    if candidate == cleaned_existing:
                        consumed_picture_values.add(index)
                        break
            else:
                for index, candidate in enumerate(source_text_values):
                    if index in consumed_text_values:
                        continue
                    if candidate == cleaned_existing:
                        consumed_text_values.add(index)
                        break
            continue

        if _is_picture_field(str(placeholder_key)):
            replacement = ""
            for index, candidate in enumerate(source_picture_values):
                if index in consumed_picture_values:
                    continue
                consumed_picture_values.add(index)
                replacement = candidate
                break
            projected[str(placeholder_key)] = replacement
            continue

        replacement = ""
        for index, candidate in enumerate(source_text_values):
            if index in consumed_text_values:
                continue
            consumed_text_values.add(index)
            replacement = candidate
            break
        projected[str(placeholder_key)] = replacement

    return projected


def _template_family(template_key: str) -> str:
    normalized = str(template_key or "").strip().lower()
    for suffix in ("_alt", "_repeat"):
        if normalized.endswith(suffix):
            return normalized[: -len(suffix)]
    return normalized


def _resolve_template_key(
    requested_template_key: str,
    *,
    slide_payload: Mapping[str, Any],
    used_template_keys: set[str],
    available_templates: Mapping[str, Any],
) -> str:
    normalized_requested = str(requested_template_key or "").strip()
    if normalized_requested in available_templates and normalized_requested not in used_template_keys:
        return normalized_requested

    payload_signature = tuple(
        sorted(
            str(key)
            for key in slide_payload.keys()
            if str(key) not in SLIDE_METADATA_FIELDS
        )
    )
    payload_has_picture = any(_is_picture_field(key) for key in payload_signature)
    requested_signature = _template_signature(available_templates.get(normalized_requested))
    requested_family = _template_family(normalized_requested)

    candidate_keys = [
        template_key
        for template_key in available_templates.keys()
        if template_key not in used_template_keys
    ]

    def candidate_rank(template_key: str) -> tuple[int, int, str]:
        template_signature = _template_signature(available_templates.get(template_key))
        template_has_picture = any(_is_picture_field(key) for key in template_signature)
        same_payload_signature = int(template_signature == payload_signature)
        same_requested_signature = int(
            bool(requested_signature) and template_signature == requested_signature
        )
        same_family = int(_template_family(template_key) == requested_family)
        picture_compatibility = int(template_has_picture == payload_has_picture)
        overlap_count = len(set(template_signature) & set(payload_signature))
        return (
            same_payload_signature,
            same_requested_signature + same_family,
            picture_compatibility,
            overlap_count,
            template_key,
        )

    candidate_keys.sort(key=lambda key: candidate_rank(key), reverse=True)
    if not candidate_keys:
        raise ValueError(f"Unknown template_key '{requested_template_key}' in slide payload")

    best_key = candidate_keys[0]
    if normalized_requested not in available_templates and best_key not in available_templates:
        raise ValueError(f"Unknown template_key '{requested_template_key}' in slide payload")
    return best_key


def canonicalize_slide_payload(
    payload: Mapping[str, Any],
    *,
    available_templates: Mapping[str, Any],
) -> dict[str, dict[str, Any]]:
    if not isinstance(payload, Mapping) or not payload:
        raise ValueError("Slide payload must be a non-empty JSON object")

    canonical: dict[str, dict[str, Any]] = {}
    used_template_keys: set[str] = set()

    for slide_index, (raw_key, raw_value) in enumerate(payload.items(), start=1):
        if not isinstance(raw_value, dict):
            raise ValueError(
                f"Slide entry '{raw_key}' must be an object, received {type(raw_value).__name__}"
            )

        slide_payload = dict(raw_value)
        if _is_slide_key(str(raw_key)):
            template_key = str(slide_payload.get(TEMPLATE_KEY_FIELD, "")).strip()
            if not template_key:
                raise ValueError(
                    f"Slide entry '{raw_key}' is missing required '{TEMPLATE_KEY_FIELD}'"
                )
        else:
            template_key = str(raw_key).strip()
            slide_payload = {TEMPLATE_KEY_FIELD: template_key, **slide_payload}

        template_key = _resolve_template_key(
            template_key,
            slide_payload=slide_payload,
            used_template_keys=used_template_keys,
            available_templates=available_templates,
        )
        slide_payload = _project_slide_payload_to_template(
            slide_payload,
            template_key=template_key,
            available_templates=available_templates,
        )
        used_template_keys.add(template_key)
        canonical[f"{SLIDE_KEY_PREFIX}{slide_index}"] = slide_payload

    return canonical


def strip_slide_payload_metadata(slide_payload: Mapping[str, Any]) -> dict[str, Any]:
    return {
        key: value
        for key, value in dict(slide_payload).items()
        if key not in SLIDE_METADATA_FIELDS
    }
