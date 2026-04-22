from __future__ import annotations

import re
from typing import Any, Mapping

from mini_layout_engine.models.layout_contract import LayoutArchetypeContract


_WS_RE = re.compile(r"\s+")


def _normalize_text(
    value: Any,
    *,
    trim_strings: bool,
    collapse_whitespace: bool,
) -> str:
    text = str(value if value is not None else "")
    if trim_strings:
        text = text.strip()
    if collapse_whitespace:
        text = _WS_RE.sub(" ", text).strip() if trim_strings else _WS_RE.sub(" ", text)
    return text


def _normalize_list_field(
    value: Any,
    *,
    trim_strings: bool,
    collapse_whitespace: bool,
    split_lines_for_lists: bool,
    coerce_single_to_list: bool,
) -> list[str]:
    if value is None:
        return []

    raw_items: list[Any] = []
    if isinstance(value, str):
        if coerce_single_to_list:
            raw_items = value.splitlines() if split_lines_for_lists else [value]
        else:
            raw_items = [value]
    elif isinstance(value, (list, tuple)):
        for item in value:
            if item is None:
                continue
            if isinstance(item, str) and split_lines_for_lists:
                raw_items.extend(item.splitlines())
            else:
                raw_items.append(item)
    else:
        if coerce_single_to_list:
            raw_items = [value]

    normalized: list[str] = []
    for item in raw_items:
        text = _normalize_text(
            item,
            trim_strings=trim_strings,
            collapse_whitespace=collapse_whitespace,
        )
        if text:
            normalized.append(text)
    return normalized


def normalize_content(
    contract: LayoutArchetypeContract,
    content: Mapping[str, Any] | None,
) -> dict[str, Any]:
    if not isinstance(content, Mapping):
        return {}

    rules = contract.normalization_rules or {}
    trim_strings = bool(rules.get("trim_strings", True))
    drop_empty_fields = bool(rules.get("drop_empty_fields", True))
    coerce_single_to_list = bool(rules.get("coerce_single_to_list", True))
    split_lines_for_lists = bool(rules.get("split_lines_for_lists", True))
    collapse_whitespace = bool(rules.get("collapse_whitespace", False))

    normalized: dict[str, Any] = {}
    content_map = {str(k): v for k, v in dict(content).items()}

    for field_name in contract.all_fields:
        if field_name not in content_map:
            continue
        field_type = contract.field_types.get(field_name, "string")
        value = content_map[field_name]

        if field_type in {"paragraph_list", "bullet_list"}:
            normalized_value = _normalize_list_field(
                value,
                trim_strings=trim_strings,
                collapse_whitespace=collapse_whitespace,
                split_lines_for_lists=split_lines_for_lists,
                coerce_single_to_list=coerce_single_to_list,
            )
            if drop_empty_fields and not normalized_value:
                continue
            normalized[field_name] = normalized_value
            continue

        normalized_text = _normalize_text(
            value,
            trim_strings=trim_strings,
            collapse_whitespace=collapse_whitespace,
        )
        if drop_empty_fields and not normalized_text:
            continue
        normalized[field_name] = normalized_text

    return normalized
