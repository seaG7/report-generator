from __future__ import annotations

from copy import deepcopy
from pathlib import Path

from docx.oxml.ns import qn
from PIL import Image

EMU_PER_INCH = 914400
A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
WP_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
A_EXT = "{http://schemas.openxmlformats.org/drawingml/2006/main}ext"


MAX_FIGURE_HEIGHT_EMU = 3_420_000  # ~9.5 cm, so intro + figure + caption fit one page


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
    extent = new_p.find(".//" + WP_EXTENT)
    if extent is not None:
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
    for ext in new_p.findall(".//" + A_EXT):
        ext.set("cx", str(cx))
        ext.set("cy", str(cy))
    new_cap = deepcopy(proto_cap)
    _set_caption_text(new_cap, template.format(n=n, caption=caption))
    body = document.element.body
    sect = body.find(qn("w:sectPr"))
    if sect is not None:
        sect.addprevious(new_p)
        sect.addprevious(new_cap)
    else:
        body.append(new_p)
        body.append(new_cap)
