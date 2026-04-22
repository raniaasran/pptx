from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import re
from typing import Iterable

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn


class MiniPptxHelper:
    def __init__(self, template_path: str | Path):
        self.template_path = str(template_path)
        self.prs = Presentation(self.template_path)

    def slide_count(self) -> int:
        return len(self.prs.slides)

    def get_slide(self, slide_index: int):
        return self.prs.slides[slide_index]

    def text_shapes(self, slide_index: int) -> list:
        slide = self.get_slide(slide_index)
        return [shape for shape in slide.shapes if getattr(shape, "has_text_frame", False)]

    def picture_shapes(self, slide_index: int) -> list:
        slide = self.get_slide(slide_index)
        return [
            shape
            for shape in slide.shapes
            if int(getattr(shape, "shape_type", 0)) == int(MSO_SHAPE_TYPE.PICTURE)
        ]

    def get_shape_by_name(self, slide_index: int, name: str):
        target_name = str(name or "").strip()
        if not target_name:
            return None
        slide = self.get_slide(slide_index)
        for shape in slide.shapes:
            if str(getattr(shape, "name", "")) == target_name:
                return shape
        return None

    def replace_text_preserve_format(self, shape, new_text: str) -> bool:
        if not getattr(shape, "has_text_frame", False):
            return False

        text_frame = shape.text_frame
        first_run = None
        if text_frame.paragraphs and text_frame.paragraphs[0].runs:
            first_run = text_frame.paragraphs[0].runs[0]

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.text = ""

        if first_run is not None:
            first_run.text = str(new_text or "")
        else:
            text_frame.text = str(new_text or "")
        return True

    def clear_text_shapes(self, slide_index: int) -> None:
        for shape in self.text_shapes(slide_index):
            self.replace_text_preserve_format(shape, "")

    def replace_picture_shape(self, shape, image_path: str | Path) -> bool:
        path = Path(image_path)
        if not path.is_file():
            return False

        left, top, width, height = shape.left, shape.top, shape.width, shape.height
        slide = shape.part.slide
        element = shape._element
        element.getparent().remove(element)
        slide.shapes.add_picture(str(path), left, top, width, height)
        return True

    def remove_shape(self, shape) -> None:
        element = shape._element
        element.getparent().remove(element)

    def delete_slide(self, slide_index: int) -> None:
        slides = list(self.prs.slides._sldIdLst)
        slide_id_to_remove = slides[slide_index].rId
        self.prs.part.drop_rel(slide_id_to_remove)
        self.prs.slides._sldIdLst.remove(slides[slide_index])

    def duplicate_slide(self, slide_index: int, insert_after_index: int | None = None):
        src_slide = self.prs.slides[slide_index]
        layout = src_slide.slide_layout
        new_slide = self.prs.slides.add_slide(layout)
        for shape in list(new_slide.shapes):
            self.remove_shape(shape)

        if insert_after_index is not None:
            sld_id_list = self.prs.slides._sldIdLst
            sld_id_elems = list(sld_id_list)
            new_elem = sld_id_elems[-1]
            sld_id_list.remove(new_elem)
            sld_id_list.insert(insert_after_index + 1, new_elem)

        rel_map: dict[str, str] = {}
        for rel_id in list(src_slide.part.rels):
            rel_obj = src_slide.part.rels[rel_id]
            try:
                reltype = str(rel_obj.reltype)
                if self._is_unsafe_copied_slide_relationship(reltype):
                    continue
                if bool(getattr(rel_obj, "is_external", False)):
                    target_ref = getattr(rel_obj, "target_ref", None)
                    if not target_ref:
                        continue
                    new_rel_id = new_slide.part.relate_to(
                        target_ref,
                        reltype,
                        is_external=True,
                    )
                else:
                    target_part = rel_obj.target_part
                    new_rel_id = new_slide.part.relate_to(target_part, reltype)
            except Exception:
                continue
            rel_map[rel_id] = new_rel_id

        for shape in src_slide.shapes:
            new_element = deepcopy(shape._element)
            for blip in new_element.findall(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
            ):
                rid = blip.get(qn("r:embed"))
                if rid and rid in rel_map:
                    blip.set(qn("r:embed"), rel_map[rid])

            for node in new_element.findall(".//*"):
                for attr_name in ("r:embed", "r:id"):
                    q_attr = qn(attr_name)
                    value = node.get(q_attr)
                    if value and value in rel_map:
                        node.set(q_attr, rel_map[value])
                    elif value:
                        node.attrib.pop(q_attr, None)
                if self._is_hyperlink_node(node) and node.get(qn("r:id")) is None:
                    parent = node.getparent()
                    if parent is not None:
                        parent.remove(node)

            new_slide.shapes._spTree.insert_element_before(new_element, "p:extLst")

        self.assert_slide_relationship_integrity(len(self.prs.slides) - 1)
        self.assert_no_duplicate_text_regions(len(self.prs.slides) - 1)
        return new_slide

    @staticmethod
    def _is_unsafe_copied_slide_relationship(reltype: str) -> bool:
        unsafe_suffixes = (
            "/notesSlide",
            "/slide",
            "/slideLayout",
        )
        return any(str(reltype).endswith(suffix) for suffix in unsafe_suffixes)

    @staticmethod
    def _is_hyperlink_node(node) -> bool:
        tag = str(getattr(node, "tag", ""))
        return tag.endswith("}hlinkClick") or tag.endswith("}hlinkHover")

    @staticmethod
    def _extract_relationship_refs(xml_text: str) -> set[str]:
        return {
            match.group(1)
            for match in re.finditer(r'r:(?:id|embed)="([^"]+)"', xml_text)
        }

    def validate_slide_relationship_integrity(self, slide_index: int) -> list[str]:
        slide = self.get_slide(slide_index)
        xml_text = slide.part.blob.decode("utf-8", errors="ignore")
        refs = self._extract_relationship_refs(xml_text)
        rel_ids = set(str(rel_id) for rel_id in list(slide.part.rels))
        missing = sorted(ref for ref in refs if ref not in rel_ids)
        return missing

    def assert_slide_relationship_integrity(self, slide_index: int) -> None:
        missing = self.validate_slide_relationship_integrity(slide_index)
        if missing:
            raise ValueError(
                f"Slide {slide_index + 1} has missing relationship IDs: {', '.join(missing)}"
            )

    def assert_all_slides_relationship_integrity(self) -> None:
        issues: list[str] = []
        for idx in range(self.slide_count()):
            missing = self.validate_slide_relationship_integrity(idx)
            if missing:
                issues.append(
                    f"slide {idx + 1}: {', '.join(missing)}"
                )
        if issues:
            raise ValueError(
                "Presentation has broken slide relationships: " + " | ".join(issues)
            )

    def validate_slide_hyperlink_integrity(self, slide_index: int) -> list[str]:
        slide = self.get_slide(slide_index)
        issues: list[str] = []
        for element in slide.part._element.iter():
            if not self._is_hyperlink_node(element):
                continue
            if element.get(qn("r:id")) is None:
                issues.append(str(element.tag))
        return issues

    def assert_all_slides_hyperlink_integrity(self) -> None:
        issues: list[str] = []
        for idx in range(self.slide_count()):
            slide_issues = self.validate_slide_hyperlink_integrity(idx)
            if slide_issues:
                issues.append(f"slide {idx + 1}: {', '.join(slide_issues)}")
        if issues:
            raise ValueError(
                "Presentation has hyperlink nodes without relationship IDs: "
                + " | ".join(issues)
            )

    def validate_slide_cross_references(self, slide_index: int) -> list[str]:
        slide = self.get_slide(slide_index)
        issues: list[str] = []
        active_slide_partnames = {
            str(self.get_slide(idx).part.partname).lstrip("/")
            for idx in range(self.slide_count())
        }
        current_slide_partname = str(slide.part.partname).lstrip("/")

        for rel_id in list(slide.part.rels):
            rel_obj = slide.part.rels[rel_id]
            reltype = str(rel_obj.reltype)
            if reltype.endswith("/slide"):
                try:
                    target_partname = str(rel_obj.target_part.partname).lstrip("/")
                except Exception:
                    target_partname = ""
                if target_partname not in active_slide_partnames:
                    issues.append(
                        f"{rel_id} points to non-active slide {target_partname or '<unknown>'}"
                    )
            elif reltype.endswith("/notesSlide"):
                try:
                    notes_part = rel_obj.target_part
                except Exception:
                    issues.append(f"{rel_id} notesSlide target is unavailable")
                    continue
                for notes_rel_id in list(notes_part.rels):
                    notes_rel = notes_part.rels[notes_rel_id]
                    if not str(notes_rel.reltype).endswith("/slide"):
                        continue
                    try:
                        back_partname = str(notes_rel.target_part.partname).lstrip("/")
                    except Exception:
                        back_partname = ""
                    if back_partname != current_slide_partname:
                        issues.append(
                            f"{rel_id} notesSlide back-reference points to "
                            f"{back_partname or '<unknown>'}"
                        )
        return issues

    def assert_all_slides_cross_references(self) -> None:
        issues: list[str] = []
        for idx in range(self.slide_count()):
            slide_issues = self.validate_slide_cross_references(idx)
            if slide_issues:
                issues.append(f"slide {idx + 1}: " + "; ".join(slide_issues))
        if issues:
            raise ValueError(
                "Presentation has invalid slide cross-references: " + " | ".join(issues)
            )

    def validate_duplicate_text_regions(self, slide_index: int) -> list[str]:
        regions: dict[tuple[int, int, int, int], list[str]] = {}
        for shape in self.text_shapes(slide_index):
            key = (int(shape.left), int(shape.top), int(shape.width), int(shape.height))
            regions.setdefault(key, []).append(str(getattr(shape, "name", "")))

        issues: list[str] = []
        for key, names in regions.items():
            if len(names) <= 1:
                continue
            issues.append(f"{key}: " + ", ".join(names))
        return issues

    def assert_no_duplicate_text_regions(self, slide_index: int) -> None:
        issues = self.validate_duplicate_text_regions(slide_index)
        if issues:
            raise ValueError(
                f"Slide {slide_index + 1} has duplicate overlapping text regions: "
                + " | ".join(issues)
            )

    def assert_all_slides_have_no_duplicate_text_regions(self) -> None:
        issues: list[str] = []
        for idx in range(self.slide_count()):
            slide_issues = self.validate_duplicate_text_regions(idx)
            if slide_issues:
                issues.append(f"slide {idx + 1}: " + " | ".join(slide_issues))
        if issues:
            raise ValueError(
                "Presentation has duplicate overlapping text regions: "
                + " || ".join(issues)
            )

    def save(self, output_path: str | Path) -> None:
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(out))

    @staticmethod
    def sort_shapes_reading_order(shapes: Iterable) -> list:
        return sorted(shapes, key=lambda shape: (int(shape.top), int(shape.left)))
