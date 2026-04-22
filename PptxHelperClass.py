from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from copy import deepcopy
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT


class PPTXHelper:
    def __init__(self, file_path: str):
        self.prs = Presentation(file_path)

    # -----------------------------
    #   SLIDE / SHAPE UTILITIES
    # -----------------------------

    def get_slide(self, number: int):
        return self.prs.slides[number]

    def get_shapes(self, slide_number: int):
        slide = self.get_slide(slide_number)
        return slide.shapes

    def get_shape_by_name(self, slide_number: int, name: str):
        slide = self.get_slide(slide_number)
        for shape in slide.shapes:
            if shape.name == name:
                return shape
        return None

    # -----------------------------
    #   TEXT HANDLING
    # -----------------------------

    def replace_text_preserve_format(self, shape, new_text: str):
        if not shape.has_text_frame:
            return False

        text_frame = shape.text_frame

        # Store first run to preserve formatting
        first_run = None
        if text_frame.paragraphs and text_frame.paragraphs[0].runs:
            first_run = text_frame.paragraphs[0].runs[0]

        # Clear existing text
        for p in text_frame.paragraphs:
            for r in p.runs:
                r.text = ""

        # Put new text
        if first_run:
            first_run.text = new_text
        else:
            text_frame.text = new_text

        return True

    def set_font_size(self, shape, size: int):
        if not shape.has_text_frame:
            return False

        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size)
        return True

    def set_font_color(self, shape, rgb_tuple):
        if not shape.has_text_frame:
            return False

        r, g, b = rgb_tuple
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(r, g, b)
        return True

    # -----------------------------
    #   SHAPE CREATION
    # -----------------------------

    def add_textbox(self, slide_number: int, left_in, top_in, width_in, height_in, text=""):
        slide = self.get_slide(slide_number)

        box = slide.shapes.add_textbox(
            Inches(left_in),
            Inches(top_in),
            Inches(width_in),
            Inches(height_in),
        )

        tf = box.text_frame
        tf.text = text
        return box

    # -----------------------------
    #   REMOVE SHAPES
    # -----------------------------

    def remove_shape(self, shape):
        sp = shape._element
        sp.getparent().remove(sp)

    # -----------------------------
    #   SAVE PRESENTATION
    # -----------------------------

    def save(self, out_path: str):
        self.prs.save(out_path)


    def copy_text_format(self, source_shape, target_shape):
        if not (source_shape.has_text_frame and target_shape.has_text_frame):
            return False

        src_tf = source_shape.text_frame
        tgt_tf = target_shape.text_frame

        # Clear default text in the new shape
        tgt_tf.clear()

        # Copy paragraph + run formatting
        for p_src in src_tf.paragraphs:
            p_tgt = tgt_tf.add_paragraph()
            
            # Copy paragraph-level formatting
            p_tgt.level = p_src.level
            p_tgt.alignment = p_src.alignment

            for r_src in p_src.runs:
                r_tgt = p_tgt.add_run()
                r_tgt.text = r_src.text

                # Copy font formatting
                font_src = r_src.font
                font_tgt = r_tgt.font

                font_tgt.size = font_src.size
                font_tgt.bold = font_src.bold
                font_tgt.italic = font_src.italic
                font_tgt.color.rgb = font_src.color.rgb if font_src.color.rgb else None
                font_tgt.name = font_src.name
                font_tgt.underline = font_src.underline

        return True

    def create_formatted_shape_from(
        self,
        slide_number,
        source_shape,
        new_text,
        left=None,
        top=None,
        width=None,
        height=None,
        offset_left_in=0,
        offset_top_in=0
    ):
        slide = self.get_slide(slide_number)

        new_left = left if left is not None else source_shape.left
        new_top = top if top is not None else source_shape.top
        new_width = width if width is not None else source_shape.width
        new_height = height if height is not None else source_shape.height

        # Apply offsets (inches → EMU)
        new_left += int(offset_left_in * 914400)
        new_top += int(offset_top_in * 914400)

        # Create new textbox
        new_shape = slide.shapes.add_textbox(new_left, new_top, new_width, new_height)

        # Copy text formatting AND insert new text using runs
        if source_shape.has_text_frame and new_shape.has_text_frame:
            src_tf = source_shape.text_frame
            tgt_tf = new_shape.text_frame
            tgt_tf.clear()

            # Split new_text into paragraphs (optional)
            new_paras = new_text.split("\n")

            for i, p_src in enumerate(src_tf.paragraphs):
                if i < len(new_paras):
                    text_to_add = new_paras[i]
                else:
                    text_to_add = ""  # if fewer lines in new_text than source

                p_tgt = tgt_tf.add_paragraph()
                p_tgt.alignment = p_src.alignment
                p_tgt.level = p_src.level

                for r_src in p_src.runs:
                    r_tgt = p_tgt.add_run()
                    r_tgt.text = text_to_add  # all new text in each run
                    font_src = r_src.font
                    font_tgt = r_tgt.font

                    font_tgt.name = font_src.name
                    font_tgt.size = font_src.size
                    font_tgt.bold = font_src.bold
                    font_tgt.italic = font_src.italic
                    font_tgt.underline = font_src.underline
                    font_tgt.color.rgb = font_src.color.rgb if font_src.color.rgb else None

        return new_shape


    
    def get_shape_position(self, shape):
        return {
            "left_emu": shape.left,
            "top_emu": shape.top,
            "width_emu": shape.width,
            "height_emu": shape.height,
            "left_in": shape.left / Inches(1),
            "top_in": shape.top / Inches(1),
            "width_in": shape.width / Inches(1),
            "height_in": shape.height / Inches(1)
        }
        
    def delete_slide(self, slide_index):
        slides = list(self.prs.slides._sldIdLst)
        slide_id_to_remove = slides[slide_index].rId
        self.prs.part.drop_rel(slide_id_to_remove)
        self.prs.slides._sldIdLst.remove(slides[slide_index])


    def duplicate_slide(self, slide_index, insert_after_index=None):
        """
        Duplicate slide at slide_index. Optionally insert after insert_after_index.
        Returns the newly created slide object.
        """
        src_slide = self.prs.slides[slide_index]
        layout = src_slide.slide_layout

        # Add new slide (appended at end)
        new_slide = self.prs.slides.add_slide(layout)

        # If user wants to insert at a specific position, move its sldId (optional)
        if insert_after_index is not None:
            # move new slide to position insert_after_index + 1
            sldIdLst = self.prs.slides._sldIdLst
            sldId_elems = list(sldIdLst)
            new_elem = sldId_elems[-1]  # the newly added slide id element
            # remove last and insert after requested position element
            sldIdLst.remove(new_elem)
            sldIdLst.insert(insert_after_index + 1, new_elem)

        # Build mapping for relationships from source slide.part to new_slide.part
        rel_map = {}
        for rel in list(src_slide.part.rels):
            rel_obj = src_slide.part.rels[rel]
            # Skip relationships that are slide-level notes or slide/layout type that shouldn't be copied
            # We'll copy most non-external relationships (images, embedded parts, charts, etc.)
            try:
                target_part = rel_obj.target_part
            except AttributeError:
                # external or unusual relationship; skip
                continue

            # create a new relationship from the new_slide.part to the same target_part
            # rel_obj.reltype contains the relationship type (string)
            try:
                new_rId = new_slide.part.relate_to(target_part, rel_obj.reltype)
            except Exception:
                # Some relationship types are not allowed; skip
                continue

            rel_map[rel] = new_rId

        # Deep copy shapes and remap r:embed attributes
        for shape in src_slide.shapes:
            el = shape._element
            new_el = deepcopy(el)

            # Remap any r:embed attributes inside the copied element to new rIds
            # Look for all blip elements (pictures reference images with r:embed)
            for blip in new_el.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                rid = blip.get(qn('r:embed'))
                if rid and rid in rel_map:
                    blip.set(qn('r:embed'), rel_map[rid])

            # Also remap any relationship attributes that might exist in other namespaces
            # common pattern: 'a:graphicData' / 'c:chart' etc use r:id or r:embed; we try a broad replacement
            for node in new_el.findall('.//*'):
                # check both 'r:embed' and 'r:id' attributes and remap if present
                for attr_name in ('r:embed', 'r:id'):
                    qattr = qn(attr_name)
                    val = node.get(qattr)
                    if val and val in rel_map:
                        node.set(qattr, rel_map[val])

            # Insert the element into new slide's spTree (before extLst to keep order)
            new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

        return new_slide



