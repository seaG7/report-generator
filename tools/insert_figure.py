from __future__ import annotations

from copy import deepcopy
from pathlib import Path

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

EMU_PER_INCH = 914400
A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
WP_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
A_EXT = "{http://schemas.openxmlformats.org/drawingml/2006/main}ext"


MAX_FIGURE_HEIGHT_EMU = 2_400_000  # ~6.7 cm starting size; leftover holes are closed by fit_figures


def set_drawing_extent(p_elem, cx: int, cy: int) -> None:
    cx, cy = max(1, int(cx)), max(1, int(cy))
    extent = p_elem.find(".//" + WP_EXTENT)
    if extent is not None:
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
    for ext in p_elem.findall(".//" + A_EXT):
        ext.set("cx", str(cx))
        ext.set("cy", str(cy))


def scale_drawing_to_height(p_elem, target_cy: int) -> bool:
    extent = p_elem.find(".//" + WP_EXTENT)
    if extent is None or extent.get("cx") is None or extent.get("cy") is None:
        return False
    cx, cy = int(extent.get("cx")), int(extent.get("cy"))
    if cy <= 0 or target_cy <= 0 or target_cy >= cy:
        return False
    set_drawing_extent(p_elem, max(1, int(cx * target_cy / cy)), target_cy)
    return True


def image_extent_emus(
    png_path: Path,
    max_width_emu: int,
    max_height_emu: int = MAX_FIGURE_HEIGHT_EMU,
) -> tuple[int, int]:
    with Image.open(png_path) as im:
        w, h = im.size
    if w <= 0:
        return max(max_width_emu, 1), max(max_height_emu, 1)
    cx = max_width_emu if max_width_emu > 0 else int(w / 96 * EMU_PER_INCH)
    cy = int(h / w * cx)
    if max_height_emu > 0 and cy > max_height_emu:
        cy = max_height_emu
        cx = int(w / h * cy)
        if max_width_emu > 0 and cx > max_width_emu:
            cx = max_width_emu
            cy = int(h / w * cx)
    return cx, cy


def add_image_rel(document, png_path: Path) -> str:
    r_id, _image = document.part.get_or_add_image(str(png_path))
    return r_id


def _style_caption_run(p_elem) -> None:
    for r in p_elem.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.insert(0, rPr)
        if rPr.find(qn("w:b")) is None:
            rPr.append(OxmlElement("w:b"))
        if rPr.find(qn("w:bCs")) is None:
            rPr.append(OxmlElement("w:bCs"))
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), "24")
        sz_cs = rPr.find(qn("w:szCs"))
        if sz_cs is None:
            sz_cs = OxmlElement("w:szCs")
            rPr.append(sz_cs)
        sz_cs.set(qn("w:val"), "24")
        fonts = rPr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rPr.append(fonts)
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            fonts.set(qn(f"w:{attr}"), "Times New Roman")


def _keep_figure_with_caption(figure_p, caption_p) -> None:
    """Glue image to its caption only — never the preceding body paragraph."""
    for p_elem, keep in ((figure_p, True), (caption_p, False)):
        pPr = p_elem.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p_elem.insert(0, pPr)
        for el in pPr.findall(qn("w:keepNext")):
            pPr.remove(el)
        if keep:
            pPr.append(OxmlElement("w:keepNext"))
    _style_caption_run(caption_p)


def _set_caption_text(p_elem, text: str) -> None:
    nodes = p_elem.findall(".//" + qn("w:t"))
    if not nodes:
        r = p_elem.makeelement(qn("w:r"), {})
        t = p_elem.makeelement(qn("w:t"), {})
        t.text = text
        r.append(t)
        p_elem.append(r)
        return
    nodes[0].text = text
    if "{http://www.w3.org/XML/1998/namespace}space" in nodes[0].attrib or text[:1].isspace() or text[-1:].isspace():
        nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in nodes[1:]:
        node.text = ""


def append_cloned_figure(
    document,
    proto_p,
    proto_cap,
    png_path: Path,
    caption: str,
    n: int,
    template: str,
    max_width_emu: int,
) -> None:
    rel_id = add_image_rel(document, png_path)
    cx, cy = image_extent_emus(png_path, max_width_emu)
    new_p = deepcopy(proto_p)
    blip = new_p.find(".//" + A_BLIP)
    if blip is None:
        raise ValueError("В прототипе рисунка нет blip")
    blip.set(R_EMBED, rel_id)
    set_drawing_extent(new_p, cx, cy)
    new_cap = deepcopy(proto_cap)
    _set_caption_text(new_cap, template.format(n=n, caption=caption))
    _keep_figure_with_caption(new_p, new_cap)
    body = document.element.body
    sect = body.find(qn("w:sectPr"))
    if sect is not None:
        sect.addprevious(new_p)
        sect.addprevious(new_cap)
    else:
        body.append(new_p)
        body.append(new_cap)
