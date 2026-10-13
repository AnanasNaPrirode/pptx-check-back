from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn

from app.parser import parse_pptx
from app.pipeline import build_report
from tests.pptx_factory import add_text, blank, presentation, save


def _open(data: bytes):
    from io import BytesIO

    from pptx import Presentation

    return Presentation(BytesIO(data))


def test_group_children_use_slide_coordinates():
    prs = presentation()
    slide = blank(prs)
    sp_tree = slide.shapes._spTree
    group = etree.SubElement(sp_tree, qn("p:grpSp"))
    nv = etree.SubElement(group, qn("p:nvGrpSpPr"))
    c_nv = etree.SubElement(nv, qn("p:cNvPr"))
    c_nv.set("id", "80")
    c_nv.set("name", "Group")
    etree.SubElement(nv, qn("p:cNvGrpSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    grp_pr = etree.SubElement(group, qn("p:grpSpPr"))
    xfrm = etree.SubElement(grp_pr, qn("a:xfrm"))
    _off_ext(xfrm, "a:off", "a:ext", 1_000_000, 500_000, 2_000_000, 1_000_000)
    _off_ext(xfrm, "a:chOff", "a:chExt", 0, 0, 2_000_000, 1_000_000)
    child = etree.SubElement(group, qn("p:sp"))
    child_nv = etree.SubElement(child, qn("p:nvSpPr"))
    child_id = etree.SubElement(child_nv, qn("p:cNvPr"))
    child_id.set("id", "81")
    child_id.set("name", "Child")
    etree.SubElement(child_nv, qn("p:cNvSpPr"))
    etree.SubElement(child_nv, qn("p:nvPr"))
    sp_pr = etree.SubElement(child, qn("p:spPr"))
    child_xfrm = etree.SubElement(sp_pr, qn("a:xfrm"))
    _off_ext(child_xfrm, "a:off", "a:ext", 100_000, 200_000, 300_000, 400_000)
    geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
    geom.set("prst", "rect")
    etree.SubElement(geom, qn("a:avLst"))
    body = etree.SubElement(child, qn("p:txBody"))
    etree.SubElement(body, qn("a:bodyPr"))
    etree.SubElement(body, qn("a:lstStyle"))
    paragraph = etree.SubElement(body, qn("a:p"))
    run = etree.SubElement(paragraph, qn("a:r"))
    r_pr = etree.SubElement(run, qn("a:rPr"))
    r_pr.set("sz", "1800")
    etree.SubElement(run, qn("a:t")).text = "В группе"

    model = parse_pptx(_open(save(prs)), [])
    elements = model.slides[0].elements
    grouped = next(element for element in elements if element.type == "group")
    child_model = next(element for element in elements if element.text == "В группе")
    assert (grouped.x, grouped.y, grouped.w, grouped.h) == (1_000_000, 500_000, 2_000_000, 1_000_000)
    assert (child_model.x, child_model.y, child_model.w, child_model.h) == (
        1_100_000,
        700_000,
        300_000,
        400_000,
    )


def test_font_size_and_color_inherited_from_layout_and_theme():
    prs = presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    title.text = "Наследование"
    for r_pr in title._element.findall(".//" + qn("a:rPr")):
        if "sz" in r_pr.attrib:
            del r_pr.attrib["sz"]
    layout_title = next(
        shape
        for shape in slide.slide_layout.shapes
        if getattr(shape, "is_placeholder", False)
        and shape.placeholder_format.type is not None
        and "TITLE" in str(shape.placeholder_format.type)
    )
    tx_body = layout_title._element.find(qn("p:txBody"))
    lst = tx_body.find(qn("a:lstStyle"))
    if lst is None:
        lst = etree.SubElement(tx_body, qn("a:lstStyle"))
    lvl = lst.find(qn("a:lvl1pPr"))
    if lvl is None:
        lvl = etree.SubElement(lst, qn("a:lvl1pPr"))
    def_rpr = lvl.find(qn("a:defRPr"))
    if def_rpr is None:
        def_rpr = etree.SubElement(lvl, qn("a:defRPr"))
    def_rpr.set("sz", "3200")

    run = title.text_frame.paragraphs[0].runs[0]
    run.font.color.theme_color = MSO_THEME_COLOR.DARK_1

    model = parse_pptx(_open(save(prs)), [])
    title_element = next(element for element in model.slides[0].elements if element.is_title)
    run_model = title_element.paragraphs[0].lines[0].runs[0]
    assert run_model.font_size_pt == 32
    assert run_model.color_rgb == (0, 0, 0)
    assert run_model.font_name


def test_font_size_falls_back_to_master_style():
    prs = presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    title.text = "Мастер"
    for r_pr in title._element.findall(".//" + qn("a:rPr")):
        if "sz" in r_pr.attrib:
            del r_pr.attrib["sz"]
    layout_title = next(
        shape
        for shape in slide.slide_layout.shapes
        if getattr(shape, "is_placeholder", False) and "TITLE" in str(shape.placeholder_format.type)
    )
    for r_pr in layout_title._element.findall(".//" + qn("a:defRPr")):
        if "sz" in r_pr.attrib:
            del r_pr.attrib["sz"]
    for r_pr in layout_title._element.findall(".//" + qn("a:rPr")):
        if "sz" in r_pr.attrib:
            del r_pr.attrib["sz"]
    master = slide.slide_layout.slide_master
    tx_styles = master.element.find(qn("p:txStyles"))
    title_style = tx_styles.find(qn("p:titleStyle"))
    lvl = title_style.find(qn("a:lvl1pPr"))
    def_rpr = lvl.find(qn("a:defRPr"))
    def_rpr.set("sz", "3000")

    model = parse_pptx(_open(save(prs)), [])
    title_element = next(element for element in model.slides[0].elements if element.is_title)
    assert title_element.paragraphs[0].lines[0].runs[0].font_size_pt == 30


def test_explicit_rgb_color_roundtrip():
    prs = presentation()
    slide = blank(prs)
    box = add_text(slide, "Цвет", 400_000, 400_000, 2_000_000, 600_000)
    run = box.text_frame.paragraphs[0].runs[0]
    run.font.color.rgb = RGBColor(10, 20, 30)
    model = parse_pptx(_open(save(prs)), [])
    element = next(item for item in model.slides[0].elements if item.text == "Цвет")
    assert element.paragraphs[0].lines[0].runs[0].color_rgb == (10, 20, 30)


def test_report_keeps_group_text_ids():
    prs = presentation()
    slide = blank(prs)
    add_text(slide, "Рядом", 400_000, 400_000, 1_000_000, 400_000)
    report = build_report("deck.pptx", None, save(prs), "[]")
    assert report["slides"][0]["elements"][0]["id"] == "s1-e1"


def _off_ext(parent, off_tag: str, ext_tag: str, x: int, y: int, w: int, h: int) -> None:
    off = etree.SubElement(parent, qn(off_tag))
    off.set("x", str(x))
    off.set("y", str(y))
    ext = etree.SubElement(parent, qn(ext_tag))
    ext.set("cx", str(w))
    ext.set("cy", str(h))
