from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping

from mini_layout_engine.engine.planning_engine import PlanningEngine
from mini_layout_engine.registry.family_registry import FamilyRegistry
from mini_layout_engine.rendering.fit_utils import (
    emu_to_inches,
    estimate_line_height_in,
    estimate_text_height_in,
    find_shrink_to_fit_font,
)
from mini_layout_engine.rendering.pptx_helper import MiniPptxHelper
from pptx.util import Pt


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

            mapped_layout = slide_entry.get("mapped_layout")
            mapped_layout = mapped_layout if isinstance(mapped_layout, Mapping) else {}
            has_placeholders = isinstance(mapped_layout.get("placeholders"), Mapping)
            content = self._select_content_source(
                slide_entry,
                prefer_input=has_placeholders,
            )

            if self._supports_paginated_table(layout_id, mapped_layout):
                rendered += self._render_paginated_table_slides(
                    helper,
                    prototype_index,
                    content,
                    mapped_layout=mapped_layout,
                    warnings=warnings,
                )
                continue

            helper.duplicate_slide(prototype_index)
            target_index = helper.slide_count() - 1

            self._render_slide_content(
                helper,
                target_index,
                content,
                mapped_layout=mapped_layout,
            )

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

        helper.assert_all_slides_relationship_integrity()
        helper.assert_all_slides_hyperlink_integrity()
        helper.assert_all_slides_cross_references()
        helper.assert_all_slides_have_no_duplicate_text_regions()
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

    def _supports_paginated_table(
        self,
        layout_id: str,
        mapped_layout: Mapping[str, Any],
    ) -> bool:
        if layout_id != "info_table":
            return False
        placeholders = mapped_layout.get("placeholders")
        if not isinstance(placeholders, Mapping):
            return False
        rows_mapping = placeholders.get("rows")
        return (
            isinstance(rows_mapping, Mapping)
            and str(rows_mapping.get("type", "")).strip() == "table"
        )

    def _render_paginated_table_slides(
        self,
        helper: MiniPptxHelper,
        prototype_index: int,
        content: Mapping[str, Any],
        *,
        mapped_layout: Mapping[str, Any],
        warnings: list[str],
    ) -> int:
        placeholders = mapped_layout.get("placeholders")
        if not isinstance(placeholders, Mapping):
            return 0

        rows_mapping = placeholders.get("rows")
        if not isinstance(rows_mapping, Mapping):
            return 0

        rows = [
            dict(row)
            for row in content.get("rows", [])
            if isinstance(row, Mapping)
        ]
        table_plan = self._build_table_pagination_plan(
            helper,
            prototype_index,
            rows_mapping,
            rows,
            warnings=warnings,
        )
        pages = table_plan["pages"]
        if not pages:
            pages = [{"rows": [], "row_heights_in": []}]

        rendered = 0
        for page_index, page in enumerate(pages):
            helper.duplicate_slide(prototype_index)
            target_index = helper.slide_count() - 1
            page_content = dict(content)
            page_content["rows"] = page["rows"]
            if page_index < len(pages) - 1:
                page_content["footer_notes"] = []

            self._render_slide_content(
                helper,
                target_index,
                page_content,
                mapped_layout=mapped_layout,
            )
            self._apply_table_page_geometry(
                helper,
                target_index,
                rows_mapping,
                page["row_heights_in"],
                float(table_plan["font_size_pt"]),
            )
            rendered += 1
        return rendered

    def _build_table_pagination_plan(
        self,
        helper: MiniPptxHelper,
        prototype_index: int,
        table_mapping: Mapping[str, Any],
        rows: list[dict[str, Any]],
        *,
        warnings: list[str],
    ) -> dict[str, Any]:
        shape = helper.get_shape_by_name(prototype_index, str(table_mapping.get("name", "")))
        if shape is None or not getattr(shape, "has_table", False):
            warnings.append("Could not paginate info_table because the table placeholder was not found.")
            return {
                "font_size_pt": 10.0,
                "pages": [{"rows": rows, "row_heights_in": []}],
            }

        table = shape.table
        start_row = int(table_mapping.get("start_row", 0) or 0)
        physical_row_count = max(1, len(table.rows) - start_row)
        table_height_in = emu_to_inches(shape.height)
        fixed_pre_start_height_in = self._table_rows_height_in(table, 0, start_row)
        body_height_in = max(0.01, table_height_in - fixed_pre_start_height_in)
        columns = table_mapping.get("columns", {})
        columns = columns if isinstance(columns, Mapping) else {}
        column_widths_in = {
            str(field): emu_to_inches(table.columns[int(column_index)].width) * 0.92
            for field, column_index in columns.items()
            if 0 <= int(column_index) < len(table.columns)
        }

        pagination = table_mapping.get("pagination", {})
        pagination = pagination if isinstance(pagination, Mapping) else {}
        template_font_pt = float(
            pagination.get("template_font_pt")
            or self._infer_table_template_font_size(table, start_row, columns)
            or 10.0
        )
        min_font_pt = float(pagination.get("min_font_pt", 7.0))
        step_pt = float(pagination.get("step_pt", 0.5))
        line_spacing = float(pagination.get("line_spacing", 1.0))
        vertical_padding_in = float(pagination.get("vertical_padding_in", 0.08))
        min_row_height_in = float(
            pagination.get(
                "min_row_height_in",
                max(0.22, estimate_line_height_in(template_font_pt) + 0.06),
            )
        )

        font_size_pt = self._choose_table_font_size(
            rows,
            column_widths_in,
            body_height_in,
            template_font_pt=template_font_pt,
            min_font_pt=min_font_pt,
            step_pt=step_pt,
            line_spacing=line_spacing,
            vertical_padding_in=vertical_padding_in,
        )
        row_heights = [
            self._estimate_table_row_height(
                row,
                column_widths_in,
                font_size_pt,
                line_spacing=line_spacing,
                vertical_padding_in=vertical_padding_in,
                min_row_height_in=min_row_height_in,
            )
            for row in rows
        ]

        pages: list[dict[str, Any]] = []
        cursor = 0
        while cursor < len(rows):
            used_height = 0.0
            page_rows: list[dict[str, Any]] = []
            page_heights: list[float] = []

            while cursor < len(rows) and len(page_rows) < physical_row_count:
                candidate_height = row_heights[cursor]
                if page_rows and used_height + candidate_height > body_height_in + 1e-6:
                    break
                page_rows.append(rows[cursor])
                page_heights.append(candidate_height)
                used_height += candidate_height
                cursor += 1

                if used_height >= body_height_in - 1e-6:
                    break

            if not page_rows:
                page_rows.append(rows[cursor])
                page_heights.append(body_height_in)
                cursor += 1

            pages.append({"rows": page_rows, "row_heights_in": page_heights})

        return {
            "font_size_pt": font_size_pt,
            "pages": pages,
        }

    def _table_rows_height_in(self, table: Any, start_row: int, end_row: int) -> float:
        bounded_start = max(0, int(start_row))
        bounded_end = min(max(0, int(end_row)), len(table.rows))
        if bounded_end <= bounded_start:
            return 0.0
        return sum(
            emu_to_inches(table.rows[row_index].height)
            for row_index in range(bounded_start, bounded_end)
        )

    def _infer_table_template_font_size(
        self,
        table: Any,
        start_row: int,
        columns: Mapping[str, Any],
    ) -> float | None:
        sizes: list[float] = []
        for row_index in range(start_row, len(table.rows)):
            for column_index in columns.values():
                try:
                    cell = table.cell(row_index, int(column_index))
                except Exception:
                    continue
                sizes.extend(self._cell_font_sizes(cell))
        if not sizes:
            return None
        return min(sizes)

    def _cell_font_sizes(self, cell: Any) -> list[float]:
        sizes: list[float] = []
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                size = getattr(run.font, "size", None)
                if size is not None:
                    sizes.append(float(size.pt))
        return sizes

    def _choose_table_font_size(
        self,
        rows: list[dict[str, Any]],
        column_widths_in: Mapping[str, float],
        table_height_in: float,
        *,
        template_font_pt: float,
        min_font_pt: float,
        step_pt: float,
        line_spacing: float,
        vertical_padding_in: float,
    ) -> float:
        chosen = float(template_font_pt)
        for row in rows:
            for field_name, width_in in column_widths_in.items():
                text = self._format_placeholder_value(row.get(field_name, ""))
                fit = find_shrink_to_fit_font(
                    text,
                    width_in=width_in,
                    height_in=table_height_in,
                    template_font_pt=template_font_pt,
                    min_font_pt=min_font_pt,
                    step_pt=step_pt,
                    line_spacing=line_spacing,
                    vertical_padding_in=vertical_padding_in,
                )
                if fit is None:
                    chosen = min(chosen, min_font_pt)
                else:
                    chosen = min(chosen, float(fit.font_size_pt))
        return chosen

    def _estimate_table_row_height(
        self,
        row: Mapping[str, Any],
        column_widths_in: Mapping[str, float],
        font_size_pt: float,
        *,
        line_spacing: float,
        vertical_padding_in: float,
        min_row_height_in: float,
    ) -> float:
        heights = [float(min_row_height_in)]
        for field_name, width_in in column_widths_in.items():
            text = self._format_placeholder_value(row.get(field_name, ""))
            heights.append(
                estimate_text_height_in(
                    text,
                    width_in=width_in,
                    font_size_pt=font_size_pt,
                    line_spacing=line_spacing,
                )
                + vertical_padding_in
            )
        return max(heights)

    def _apply_table_page_geometry(
        self,
        helper: MiniPptxHelper,
        slide_index: int,
        table_mapping: Mapping[str, Any],
        row_heights_in: list[float],
        font_size_pt: float,
    ) -> None:
        shape = helper.get_shape_by_name(slide_index, str(table_mapping.get("name", "")))
        if shape is None or not getattr(shape, "has_table", False):
            return

        table = shape.table
        start_row = int(table_mapping.get("start_row", 0) or 0)
        physical_row_count = max(1, len(table.rows) - start_row)
        total_table_height_in = emu_to_inches(shape.height)
        fixed_pre_start_height_in = self._table_rows_height_in(table, 0, start_row)
        body_height_in = max(0.01, total_table_height_in - fixed_pre_start_height_in)
        shrink_underfilled = len(row_heights_in) < physical_row_count
        allocated = self._allocate_table_row_heights(
            row_heights_in,
            physical_row_count,
            body_height_in,
            shrink_underfilled=shrink_underfilled,
        )

        for offset, row_height_in in enumerate(allocated):
            row_index = start_row + offset
            if row_index >= len(table.rows):
                break
            table.rows[row_index].height = int(round(row_height_in * 914400))
            for column_index in range(len(table.columns)):
                cell = table.cell(row_index, column_index)
                self._set_cell_font_size(cell, font_size_pt)

        if shrink_underfilled:
            actual_table_height_in = fixed_pre_start_height_in + sum(allocated)
            shape.height = int(round(max(0.01, actual_table_height_in) * 914400))

    def _allocate_table_row_heights(
        self,
        used_row_heights_in: list[float],
        physical_row_count: int,
        fixed_height_in: float,
        *,
        shrink_underfilled: bool = False,
    ) -> list[float]:
        used_count = min(len(used_row_heights_in), physical_row_count)
        blank_count = max(0, physical_row_count - used_count)
        if shrink_underfilled:
            blank_height_in = 0.0
        elif used_count:
            blank_height_in = 0.01
        else:
            blank_height_in = fixed_height_in / physical_row_count
        available_for_used = max(
            0.0,
            fixed_height_in - (blank_count * blank_height_in),
        )

        if used_count == 0:
            return [blank_height_in for _ in range(physical_row_count)]

        used = [max(0.01, float(height)) for height in used_row_heights_in[:used_count]]
        if shrink_underfilled:
            return used + [blank_height_in for _ in range(blank_count)]

        required = sum(used)
        if required <= 0:
            allocated_used = [available_for_used / used_count for _ in range(used_count)]
        elif required <= available_for_used:
            extra = (available_for_used - required) / used_count
            allocated_used = [height + extra for height in used]
        else:
            scale = available_for_used / required
            allocated_used = [max(0.01, height * scale) for height in used]

        return allocated_used + [blank_height_in for _ in range(blank_count)]

    def _set_cell_font_size(self, cell: Any, font_size_pt: float) -> None:
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(font_size_pt)

    def _select_content_source(
        self,
        slide_entry: Mapping[str, Any],
        *,
        prefer_input: bool = False,
    ) -> dict[str, Any]:
        keys = (
            ("input_content", "normalized_content", "validated_content", "content")
            if prefer_input
            else ("normalized_content", "validated_content", "input_content", "content")
        )
        for key in keys:
            value = slide_entry.get(key)
            if isinstance(value, Mapping):
                return {str(k): v for k, v in dict(value).items()}
        return {}

    def _render_slide_content(
        self,
        helper: MiniPptxHelper,
        slide_index: int,
        content: Mapping[str, Any],
        *,
        mapped_layout: Mapping[str, Any] | None = None,
    ) -> None:
        placeholders = {}
        if isinstance(mapped_layout, Mapping):
            raw_placeholders = mapped_layout.get("placeholders")
            if isinstance(raw_placeholders, Mapping):
                placeholders = {str(k): v for k, v in dict(raw_placeholders).items()}

        if placeholders and self._render_placeholders(
            helper,
            slide_index,
            content,
            placeholders,
        ):
            return

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

    def _render_placeholders(
        self,
        helper: MiniPptxHelper,
        slide_index: int,
        content: Mapping[str, Any],
        placeholders: Mapping[str, Any],
    ) -> bool:
        rendered_any = False
        for field, mapping in placeholders.items():
            if not isinstance(mapping, Mapping):
                continue
            field_name = str(field)
            value = content.get(field_name)
            mapping_type = str(mapping.get("type", "")).strip()

            if mapping_type == "table":
                rendered_any = self._render_table_placeholder(
                    helper,
                    slide_index,
                    mapping,
                    value,
                ) or rendered_any
                continue

            if mapping_type == "repeated_group":
                rendered_any = self._render_repeated_group_placeholder(
                    helper,
                    slide_index,
                    mapping,
                    value,
                ) or rendered_any
                continue

            if "names" in mapping:
                rendered_any = self._render_multi_shape_placeholder(
                    helper,
                    slide_index,
                    mapping,
                    value,
                ) or rendered_any
                continue

            shape_name = str(mapping.get("name", "")).strip()
            shape = helper.get_shape_by_name(slide_index, shape_name)
            if shape is None:
                continue
            helper.replace_text_preserve_format(shape, self._format_placeholder_value(value))
            rendered_any = True
        return rendered_any

    def _render_multi_shape_placeholder(
        self,
        helper: MiniPptxHelper,
        slide_index: int,
        mapping: Mapping[str, Any],
        value: Any,
    ) -> bool:
        names = mapping.get("names")
        if not isinstance(names, list):
            return False
        values = self._as_list(value)
        rendered_any = False
        for idx, name in enumerate(names):
            shape = helper.get_shape_by_name(slide_index, str(name))
            if shape is None:
                continue
            text = values[idx] if idx < len(values) else ""
            helper.replace_text_preserve_format(shape, text)
            rendered_any = True
        return rendered_any

    def _render_table_placeholder(
        self,
        helper: MiniPptxHelper,
        slide_index: int,
        mapping: Mapping[str, Any],
        value: Any,
    ) -> bool:
        shape = helper.get_shape_by_name(slide_index, str(mapping.get("name", "")))
        if shape is None or not getattr(shape, "has_table", False):
            return False
        rows = value if isinstance(value, list) else []
        table = shape.table
        start_row = int(mapping.get("start_row", 0) or 0)
        columns = mapping.get("columns", {})
        columns = columns if isinstance(columns, Mapping) else {}

        for row_index in range(start_row, len(table.rows)):
            item_index = row_index - start_row
            item = rows[item_index] if item_index < len(rows) and isinstance(rows[item_index], Mapping) else {}
            for field_name, column_index in columns.items():
                try:
                    cell = table.cell(row_index, int(column_index))
                except Exception:
                    continue
                cell.text = self._format_placeholder_value(item.get(str(field_name), ""))
        return True

    def _render_repeated_group_placeholder(
        self,
        helper: MiniPptxHelper,
        slide_index: int,
        mapping: Mapping[str, Any],
        value: Any,
    ) -> bool:
        item_mappings = mapping.get("items")
        if not isinstance(item_mappings, list):
            return False
        values = value if isinstance(value, list) else []
        rendered_any = False
        for idx, item_mapping in enumerate(item_mappings):
            if not isinstance(item_mapping, Mapping):
                continue
            item = values[idx] if idx < len(values) and isinstance(values[idx], Mapping) else {}
            for field_name, shape_name in item_mapping.items():
                shape = helper.get_shape_by_name(slide_index, str(shape_name))
                if shape is None:
                    continue
                helper.replace_text_preserve_format(
                    shape,
                    self._format_placeholder_value(item.get(str(field_name), "")),
                )
                rendered_any = True
        return rendered_any

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

    def _format_placeholder_value(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, list):
            return "\n".join(self._as_list(value))
        if isinstance(value, Mapping):
            return "\n".join(
                self._as_text(v)
                for v in value.values()
                if self._as_text(v)
            )
        return self._as_text(value)


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
