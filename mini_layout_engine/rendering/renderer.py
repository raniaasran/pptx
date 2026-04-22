from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping

from mini_layout_engine.engine.planning_engine import PlanningEngine
from mini_layout_engine.registry.family_registry import FamilyRegistry
from mini_layout_engine.rendering.pptx_helper import MiniPptxHelper


@dataclass
class RenderResult:
    output_path: str
    rendered_slides: int
    skipped_slides: int
    warnings: list[str]

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": self.output_path,
            "rendered_slides": self.rendered_slides,
            "skipped_slides": self.skipped_slides,
            "warnings": list(self.warnings),
        }


class MiniPptxRenderer:
    def __init__(self, templates_dir: str | Path | None = None):
        package_root = Path(__file__).resolve().parents[1]
        self.templates_dir = Path(templates_dir or (package_root / "templates")).resolve()
        self.family_registry = FamilyRegistry.from_specs_dir(
            package_root / "specs" / "families"
        )

    def render_from_request(
        self,
        request: Mapping[str, Any] | dict[str, Any],
        *,
        output_path: str | Path,
    ) -> RenderResult:
        planning_engine = PlanningEngine.from_default_specs()
        plan_result = planning_engine.plan_deck(request).to_dict()
        return self.render_from_plan(plan_result, output_path=output_path)

    def render_from_plan(
        self,
        plan: Mapping[str, Any] | dict[str, Any],
        *,
        output_path: str | Path,
    ) -> RenderResult:
        payload = dict(plan)
        family_id = str(payload.get("design_family_id", "")).strip()
        if not family_id:
            raise ValueError("Plan payload is missing 'design_family_id'.")

        family = self.family_registry.require(family_id)
        template_name = str(payload.get("template_file") or family.template_file).strip()
        if not template_name:
            raise ValueError(f"Family '{family_id}' has no template_file configured.")

        template_path = self.templates_dir / template_name
        if not template_path.is_file():
            raise FileNotFoundError(f"Template file not found: {template_path}")

        helper = MiniPptxHelper(template_path)
        slides_payload = payload.get("slides", [])
        if not isinstance(slides_payload, list):
            raise ValueError("Plan payload field 'slides' must be a list.")

        warnings: list[str] = []
        prototype_map = self._resolve_prototype_indices(
            family_layout_mapping=family.layout_mapping,
            total_template_slides=helper.slide_count(),
        )

        original_slide_count = helper.slide_count()
        rendered = 0
        skipped = 0

        for slide_entry in slides_payload:
            if not isinstance(slide_entry, Mapping):
                skipped += 1
                warnings.append("Skipped invalid slide entry (not an object).")
                continue

            readiness = str(slide_entry.get("readiness", "")).strip().lower()
            if readiness and readiness != "ready":
                skipped += 1
                warnings.append(
                    f"Skipped slide with layout '{slide_entry.get('layout_id')}' due to readiness='{readiness}'."
                )
                continue

            layout_id = str(slide_entry.get("layout_id", "")).strip()
            if not layout_id:
                skipped += 1
                warnings.append("Skipped slide without layout_id.")
                continue

            prototype_index = prototype_map.get(layout_id)
            if prototype_index is None:
                skipped += 1
                warnings.append(
                    f"Skipped slide with layout '{layout_id}' because family mapping has no prototype index."
                )
                continue

            helper.duplicate_slide(prototype_index)
            target_index = helper.slide_count() - 1

            content = self._select_content_source(slide_entry)
            self._render_slide_content(helper, target_index, content)

            image_ref = self._extract_image_ref(content)
            if image_ref:
                image_path = self._resolve_image_path(image_ref)
                if image_path is None:
                    warnings.append(
                        f"Image '{image_ref}' not found for layout '{layout_id}'; kept template image."
                    )
                else:
                    picture_shapes = helper.sort_shapes_reading_order(helper.picture_shapes(target_index))
                    if picture_shapes:
                        helper.replace_picture_shape(picture_shapes[0], image_path)
                    else:
                        warnings.append(
                            f"No picture shape available for image '{image_ref}' in layout '{layout_id}'."
                        )

            rendered += 1

        for index in reversed(range(original_slide_count)):
            helper.delete_slide(index)

        out_path = Path(output_path).resolve()
        helper.save(out_path)
        return RenderResult(
            output_path=str(out_path),
            rendered_slides=rendered,
            skipped_slides=skipped,
            warnings=warnings,
        )

    def _resolve_prototype_indices(
        self,
        *,
        family_layout_mapping: Mapping[str, Mapping[str, Any]],
        total_template_slides: int,
    ) -> dict[str, int]:
        prototype_map: dict[str, int] = {}
        fallback_index = 0
        for layout_id, mapping in family_layout_mapping.items():
            if not isinstance(mapping, Mapping):
                continue
            explicit = mapping.get("template_slide_index")
            if explicit is not None:
                try:
                    resolved = int(explicit) - 1
                except Exception:
                    resolved = -1
            else:
                resolved = fallback_index
                fallback_index += 1

            if 0 <= resolved < total_template_slides:
                prototype_map[str(layout_id)] = resolved
        return prototype_map

    def _select_content_source(self, slide_entry: Mapping[str, Any]) -> dict[str, Any]:
        for key in ("normalized_content", "validated_content", "input_content", "content"):
            value = slide_entry.get(key)
            if isinstance(value, Mapping):
                return {str(k): v for k, v in dict(value).items()}
        return {}

    def _render_slide_content(
        self,
        helper: MiniPptxHelper,
        slide_index: int,
        content: Mapping[str, Any],
    ) -> None:
        text_shapes = helper.sort_shapes_reading_order(helper.text_shapes(slide_index))
        if not text_shapes:
            return

        blocks = self._build_text_blocks(content)
        if not blocks:
            helper.clear_text_shapes(slide_index)
            return

        helper.replace_text_preserve_format(text_shapes[0], blocks[0])
        remaining_blocks = blocks[1:]
        if not remaining_blocks:
            for shape in text_shapes[1:]:
                helper.replace_text_preserve_format(shape, "")
            return

        target_shapes = text_shapes[1:] or text_shapes[:1]
        for idx, shape in enumerate(target_shapes):
            if idx < len(remaining_blocks):
                helper.replace_text_preserve_format(shape, remaining_blocks[idx])
            elif idx == len(target_shapes) - 1 and len(remaining_blocks) > len(target_shapes):
                overflow_start = len(target_shapes) - 1
                helper.replace_text_preserve_format(
                    shape,
                    "\n\n".join(remaining_blocks[overflow_start:]),
                )
            else:
                helper.replace_text_preserve_format(shape, "")

    def _build_text_blocks(self, content: Mapping[str, Any]) -> list[str]:
        blocks: list[str] = []
        used_fields: set[str] = set()

        title = self._as_text(content.get("title"))
        if title:
            blocks.append(title)
            used_fields.add("title")

        for key in ("intro_paragraphs",):
            values = self._as_list(content.get(key))
            if values:
                blocks.append("\n".join(values))
                used_fields.add(key)

        pair_definitions = [
            ("card_titles", "card_descriptions"),
            ("point_titles", "point_descriptions"),
            ("step_titles", "step_descriptions"),
        ]
        for titles_key, desc_key in pair_definitions:
            titles = self._as_list(content.get(titles_key))
            descs = self._as_list(content.get(desc_key))
            if not titles and not descs:
                continue
            used_fields.update({titles_key, desc_key})
            max_len = max(len(titles), len(descs))
            for idx in range(max_len):
                title_text = titles[idx] if idx < len(titles) else ""
                desc_text = descs[idx] if idx < len(descs) else ""
                if title_text and desc_text:
                    blocks.append(f"{idx + 1}. {title_text}\n{desc_text}")
                elif title_text:
                    blocks.append(f"{idx + 1}. {title_text}")
                elif desc_text:
                    blocks.append(f"{idx + 1}. {desc_text}")

        for key, value in content.items():
            field = str(key)
            if field in used_fields or field == "image":
                continue
            if isinstance(value, list):
                values = self._as_list(value)
                if values:
                    blocks.append("\n".join(values))
                continue
            text = self._as_text(value)
            if text:
                blocks.append(text)

        return blocks

    def _extract_image_ref(self, content: Mapping[str, Any]) -> str | None:
        image_value = content.get("image")
        image_text = self._as_text(image_value)
        return image_text or None

    def _resolve_image_path(self, image_ref: str) -> Path | None:
        ref = str(image_ref).strip()
        if not ref:
            return None

        raw = Path(ref)
        if raw.is_file():
            return raw.resolve()

        package_root = Path(__file__).resolve().parents[1]
        search_dirs = [
            package_root / "examples" / "media",
            package_root / "examples" / "images",
            Path.cwd() / "data" / "1" / "images",
            Path.cwd() / "data" / "1" / "media",
        ]

        candidate_names = [raw.name]
        if raw.suffix == "":
            for ext in (".png", ".jpg", ".jpeg", ".webp"):
                candidate_names.append(f"{raw.name}{ext}")

        for directory in search_dirs:
            if not directory.exists():
                continue
            for name in candidate_names:
                path = directory / name
                if path.is_file():
                    return path.resolve()
        return None

    @staticmethod
    def _as_text(value: Any) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        return text

    @staticmethod
    def _as_list(value: Any) -> list[str]:
        if value is None:
            return []
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        text = str(value).strip()
        if not text:
            return []
        return [part.strip() for part in text.splitlines() if part.strip()]


def load_json_payload(path: str | Path) -> Any:
    file_path = Path(path)
    return json.loads(file_path.read_text(encoding="utf-8"))


def normalize_request_payload(payload: Any) -> dict[str, Any]:
    if isinstance(payload, list):
        if not payload:
            raise ValueError("Request payload list is empty.")
        first = payload[0]
        if not isinstance(first, Mapping):
            raise ValueError("First request item must be an object.")
        return dict(first)
    if isinstance(payload, Mapping):
        return dict(payload)
    raise ValueError("Request payload must be an object or list of objects.")
