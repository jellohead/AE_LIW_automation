# update_paragraphs.py
# Adds list items to a textbox on a PowerPoint slide
"""
Update the text in a textbox on a PowerPoint slide with a list of strings.

The font size of each paragraph is determined by whether the string starts with a
'+' or a '-' (or neither). The font sizes can be specified, but the default is
24pt for 'neutral' paragraphs, 20pt for '+' paragraphs, and 18pt for '-' paragraphs.

The paragraphs are left-aligned.

The strings are stripped of leading and trailing whitespace, and any leading '+' or
'-' is removed from the string before adding it to the textbox.

The first paragraph is added to the existing first paragraph in the textbox, if
it is empty. If it is not empty, the new paragraph is added after it. Subsequent
paragraphs are added after the previous one.

To generate a visually blank line between paragraph strings, use ' ' rather than '' or '\n' when creating the
paragraph_strings variable in the calling function.
    1. python-pptx will flatten the new line characters and not treat them as a place to start a new line.
    2. Using '' will result in "IndexError: tuple index out of range" since the run will be empty.

The suggested use case is when you want to dynamically generate some bullet
points based on some data. For example, you might want to report on which
products are currently in stock, and which are not.

"""
from openpyxl.pivot.fields import Boolean
from pptx.util import Pt, Inches
import re
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from typing import List, Optional
from pptx.text.text import TextFrame
from AE_LIW_automation.helper_modules.format_paragraph_xml import format_paragraph_xml
# from src.Mahindra_ROXOR_CSAT_automation.helper_modules import update_paragraphs

def update_paragraphs(paragraph_strings: List[str],
                      text_holder: TextFrame,
                      font_name: str = 'Tahoma',
                      # level 0 settings
                      l0_font_color: tuple = (0, 0, 0), # black
                      l0_font_size: Pt = Pt(24),
                      l0_font_bold: bool = False,
                      l0_font_italic: bool = False,
                      l0_h_alignment = PP_ALIGN.LEFT,
                      l0_v_alignment = MSO_ANCHOR.MIDDLE,
                      l0_line_spacing: Pt = Pt(20),
                      l0_left_indent: Optional[str] = None,
                      l0_hanging_indent: Optional[float] = -0.1,
                      l0_bullet_char: str = u'\u2022', # \u2022 (•)
                      # level 1 settings
                      l1_font_color: tuple = (0, 0, 0), # black
                      l1_font_size: Pt = Pt(20),
                      l1_font_bold: bool = False,
                      l1_font_italic: bool = False,
                      l1_h_alignment = PP_ALIGN.LEFT,
                      l1_v_alignment = MSO_ANCHOR.MIDDLE,
                      l1_line_spacing: Pt = Pt(10),
                      l1_left_indent: Optional[str] = None,
                      l1_hanging_indent: Optional[float] = -0.1,
                      l1_bullet_char: str = u'\u2022',
                      # level 2 settings
                      l2_font_color = (0, 0, 0), # black
                      l2_font_size: Pt = Pt(18),
                      l2_font_bold: bool = False,
                      l2_font_italic: bool = False,
                      l2_h_alignment = PP_ALIGN.LEFT,
                      l2_v_alignment = MSO_ANCHOR.MIDDLE,
                      l2_line_spacing = Pt(10),
                      l2_left_indent: Optional[float] = 0.5,
                      l2_hanging_indent: Optional[float] = -0.1,
                      l2_bullet_char: str = u'\u2022',
                      ) -> None:

    # set textbox fill before adding paragraphs
    # shape = text_holder._parent
    # shape.fill.solid()
    # shape.fill.fore_color.rgb = RGBColor(230, 230, 250)

    for i, para_string in enumerate(paragraph_strings):
        clean_text = re.sub(r'^[+-]', '', para_string).strip()

        # check if paragraph is empty to prevent adding an extra line at the head of the text block
        if i == 0 and text_holder.text == "":
            p = text_holder.paragraphs[0]
        else:
            p = text_holder.add_paragraph()

        p.text = clean_text
        run = p.runs[0]
        run.font.name = font_name

        # handle different bullet levels
        if para_string.startswith('+'):
            level = 1
            run.font.color.rgb = RGBColor(*l1_font_color) # black is default
            run.font.size = l1_font_size
            run.font.bold = l1_font_bold
            run.font.italic = l1_font_italic
            p.alignment = l1_h_alignment
            p.line_spacing = l1_line_spacing
            format_paragraph_xml(p,
                                 level=level,
                                 left_indent=l1_left_indent,
                                 # hanging_indent=l1_hanging_indent,
                                 # bullet_char=l1_bullet_char
                                 )
            # p.paragraph_format.left_indent = l1_indent
        elif para_string.startswith('-'):
            level = 2
            # if l2_indent is not None:
            #     pf.format.left_indent = l1_indent
            run.font.color.rgb = RGBColor(*l2_font_color)
            run.font.size = l2_font_size
            run.font.bold = l2_font_bold
            run.font.italic = l2_font_italic
            p.alignment = l2_h_alignment
            p.line_spacing = l2_line_spacing
            format_paragraph_xml(p,
                                 level=level,
                                 left_indent=l2_left_indent,
                                 # hanging_indent=l2_hanging_indent,
                                 # bullet_char=l2_bullet_char
                                 )
        else:
            level = 0
            # if l0_indent is not None:
            #     pf.left_indent = l1_indent
            run.font.color.rgb = RGBColor(*l0_font_color)
            run.font.size = l0_font_size
            run.font.bold = l0_font_bold
            run.font.italic = l0_font_italic
            p.alignment = l0_h_alignment
            p.line_spacing = l0_line_spacing
            format_paragraph_xml(p,
                                 level=level,
                                 left_indent=l0_left_indent,
                                 # hanging_indent=l0_hanging_indent,
                                 # bullet_char=l0_bullet_char
            )
