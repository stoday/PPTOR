"""
Utility helpers for common PowerPoint edits using python-pptx.
These wrap risky/verbose operations into safe, tested functions to reduce model hallucinations.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR, MSO_SHAPE_TYPE
from pptx.oxml.xmlchemy import OxmlElement


# -------- Basic getters --------
def load_presentation(ppt_path: str) -> Presentation:
    """Load presentation from path."""
    return Presentation(ppt_path)


def get_slide(prs: Presentation, index: int):
    """Safe slide access with 0-based index."""
    if index < 0 or index >= len(prs.slides):
        raise IndexError(f"Slide index {index} out of range; total slides: {len(prs.slides)}")
    return prs.slides[index]


# -------- Shape utilities --------
def delete_shapes_except(slide, shapes_to_keep):
    """Delete all shapes except those in shapes_to_keep."""
    keep_elements = {s.element for s in shapes_to_keep if s is not None}
    for shape in list(slide.shapes):
        if shape.element in keep_elements:
            continue
        sp = shape.element
        sp.getparent().remove(sp)


def remove_connectors_and_lines(slide):
    """Remove existing connector/line shapes."""
    connector_type = getattr(MSO_SHAPE_TYPE, "CONNECTOR", None)
    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.LINE or (connector_type and shape.shape_type == connector_type):
            sp = shape.element
            sp.getparent().remove(sp)


def add_rounded_textbox(
    slide,
    text: str,
    left,
    top,
    width,
    height,
    fill_rgb=(232, 244, 248),
    text_rgb=(50, 50, 50),
    font_size=20,
    align=PP_ALIGN.CENTER,
    vertical_anchor=MSO_ANCHOR.MIDDLE,
):
    """Add a rounded rectangle with text and return the shape."""
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)

    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*fill_rgb)

    line = shape.line
    line.color.rgb = RGBColor(180, 180, 180)
    line.width = Pt(1)

    if hasattr(shape, "adjustments") and len(shape.adjustments) > 0:
        shape.adjustments[0] = 0.15  # rounded corner

    # Shadow (color not exposed in python-pptx; leave default)
    shadow = shape.shadow
    shadow.inherit = False
    shadow.style = "OUTER"
    shadow.distance = Pt(3)
    shadow.angle = 45
    shadow.blur_radius = Pt(4)
    shadow.transparency = 0.5

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = vertical_anchor
    tf.margin_left = tf.margin_right = Inches(0.1)
    tf.margin_top = tf.margin_bottom = Inches(0.1)

    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Microsoft JhengHei"
    p.font.size = Pt(font_size)
    p.font.color.rgb = RGBColor(*text_rgb)
    p.alignment = align

    return shape


def add_arrow_between(slide, shape_from, shape_to, color_rgb=(70, 70, 70), width_pt=2.5, arrow_head=2):
    """Add straight connector arrow between two shapes and return the connector."""
    start_x = Emu(shape_from.left + shape_from.width)
    start_y = Emu(shape_from.top + shape_from.height // 2)
    end_x = Emu(shape_to.left)
    end_y = Emu(shape_to.top + shape_to.height // 2)

    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, start_x, start_y, end_x, end_y)
    connector.line.width = Pt(width_pt)
    connector.line.color.rgb = RGBColor(*color_rgb)
    # Apply a visible arrowhead via XML since python-pptx lacks arrow enums in older versions
    ln = connector._element.xpath(".//a:ln")
    if ln:
        tail_end = OxmlElement("a:tailEnd")
        tail_end.set("type", "triangle")
        ln[0].append(tail_end)
    return connector


# -------- Layout helpers --------
def distribute_horizontally(slide_width, count, box_width, gap, margin=Inches(0.5)):
    """
    Compute left positions to distribute `count` boxes horizontally.
    Returns list of left positions (Emu).
    """
    total_width = count * box_width + (count - 1) * gap
    start_left = (slide_width - total_width) // 2
    return [start_left + i * (box_width + gap) for i in range(count)]
