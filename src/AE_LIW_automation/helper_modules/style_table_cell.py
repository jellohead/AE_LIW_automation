from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt

# TODO: v_alignment is not working

def style_table_cell(cell, text=None,
                     font_size=12, bold=False,
                     color=RGBColor(0, 0, 0),
                     bg_color=None,
                     h_alignment=PP_ALIGN.CENTER,
                     v_alignment=MSO_ANCHOR.MIDDLE,
                     ):
    """
    Applies font size, boldness, color, alignment, and background to a table cell.

    Parameters:
        cell      : PowerPoint table cell object
        text      : Optional; if provided, sets the cell text
        font_size : Font size in points
        bold      : Boolean, whether font is bold
        color     : RGBColor instance for font color
        bg_color  : Optional RGBColor for background fill
        align     : Paragraph alignment (e.g., PP_ALIGN.CENTER)
    """
    if text is not None:
        cell.text = text


    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = h_alignment

    if paragraph.runs:
        run = paragraph.runs[0]
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color

    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color

    # this is supposed to align the text vertically in the cell but it is not working
    cell.text_frame.vertical_anchor = v_alignment
