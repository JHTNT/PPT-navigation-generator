from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from ppt_nav.outline import Outline, OutlineItem, SlidePlanEntry


class PresentationBuilder:
    def __init__(self, font_size: Optional[float] = None) -> None:
        self.nav_side_margin = Inches(0)
        self.nav_top_margin = Inches(0)
        self.nav_separator_width = Inches(0.04)
        self.body_margin_top = Inches(0.4)
        self.body_side_margin = Inches(0.5)
        self.inactive_color = RGBColor(200, 200, 200)
        self.active_color = RGBColor(0, 0, 0)
        # Base font size in points for navigation and body
        self.font_size_pt = font_size if font_size and font_size > 0 else 28.0
        # Scale navigation row height with font size (28pt -> 0.5" baseline)
        base_nav_height_in = 0.5
        scale = self.font_size_pt / 28.0
        self.nav_row_height = Inches(base_nav_height_in * scale)
        self._target_width = Inches(16)
        self._target_height = Inches(9)
        self._slide_width = 0
        self._slide_height = 0

    def build(self, outline: Outline, output_path: Path) -> None:
        prs = Presentation()
        prs.slide_width = self._target_width
        prs.slide_height = self._target_height
        self._slide_width = prs.slide_width
        self._slide_height = prs.slide_height
        for plan_entry in outline.iter_slide_plan():
            self._add_slide(prs, outline.sections, plan_entry)
        prs.save(output_path)

    def _add_slide(
        self,
        prs: Presentation,
        sections: Iterable[OutlineItem],
        plan_entry: SlidePlanEntry,
    ) -> None:
        layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(layout)
        nav_bottom = self._add_navigation(slide, sections, plan_entry)
        self._add_body_placeholder(slide, plan_entry, nav_bottom)

    def _add_navigation(
        self,
        slide,
        sections: Iterable[OutlineItem],
        plan_entry: SlidePlanEntry,
    ) -> float:
        current_section = plan_entry.section
        current_child = plan_entry.child
        top = self.nav_top_margin
        section_titles = [section.title for section in sections]
        top = self._draw_nav_row(slide, section_titles, current_section.title, top)
        top = self._draw_separator(slide, top)
        if current_section.children:
            child_titles = [child.title for child in current_section.children]
            active_child = current_child.title if current_child else None
            top = self._draw_nav_row(slide, child_titles, active_child, top)
            top = self._draw_separator(slide, top)
        return top

    def _draw_nav_row(self, slide, titles, active_title: Optional[str], top: float) -> float:
        if not titles:
            return top
        count = len(titles)
        usable_width = self._slide_width - self.nav_side_margin * 2
        # Keep a constant padding share per tab so whitespace feels uniform, then add
        # extra width based on character count to honor the "size by words" request.
        char_counts = [max(len(title.strip()) or len(title), 1) for title in titles]
        avg_chars = sum(char_counts) / count
        base_weight = max(int(avg_chars * 0.5), 4)
        tab_weights = [base_weight + chars for chars in char_counts]
        total_weight = sum(tab_weights)
        left = self.nav_side_margin
        allocated_width = 0.0
        for idx, title in enumerate(titles):
            weight = tab_weights[idx]
            if idx == count - 1:
                item_width = usable_width - allocated_width
            else:
                item_width = usable_width * (weight / total_weight)
                allocated_width += item_width
            box = slide.shapes.add_textbox(left, top, item_width, self.nav_row_height)
            tf = box.text_frame
            tf.clear()
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = tf.margin_right = 0
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER
            run = para.add_run()
            run.text = title
            run.font.size = Pt(self.font_size_pt)
            if title == active_title:
                run.font.bold = True
                run.font.color.rgb = self.active_color
            else:
                run.font.bold = True
                run.font.color.rgb = self.inactive_color
            left += item_width
            if idx < count - 1:
                self._draw_vertical_separator(slide, left, top, self.nav_row_height)
        return top + self.nav_row_height

    def _draw_separator(self, slide, top: float) -> float:
        shape = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            self.nav_side_margin,
            top,
            self._slide_width - self.nav_side_margin * 2,
            self.nav_separator_width,
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.inactive_color
        shape.line.fill.background()
        shape.shadow.inherit = False
        return top + self.nav_separator_width

    def _draw_vertical_separator(self, slide, x_pos: float, top: float, height: float) -> None:
        shape = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            x_pos - self.nav_separator_width / 2,
            top,
            self.nav_separator_width,
            height,
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.inactive_color
        shape.line.fill.background()
        shape.shadow.inherit = False

    def _add_body_placeholder(
        self,
        slide,
        plan_entry: SlidePlanEntry,
        nav_bottom: float,
    ) -> None:
        body_top = nav_bottom + self.body_margin_top
        height = self._slide_height - body_top - Inches(0.5)
        width = self._slide_width - self.body_side_margin * 2
        box = slide.shapes.add_textbox(self.body_side_margin, body_top, width, height)
        tf = box.text_frame
        tf.clear()
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT
        tf.paragraphs[0].space_after = 0
        tf.paragraphs[0].space_before = 0
        tf.paragraphs[0].line_spacing = 1.1
        para.font.size = Pt(self.font_size_pt)
        para.text = ""
        tf.add_paragraph()
