from __future__ import annotations

import argparse
import re
import sys
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from tools.config import TitleFields, load_title_fields
from tools.insert_figure import append_cloned_figure
from tools.parse_report import parse_report_md
from tools.paths import assert_not_sample
from tools.screenshot import load_manifest
from tools.toc import insert_word_toc


class ComposeError(Exception):
    pass


@dataclass
class Prototypes:
    figure_p: object | None
    figure_caption_p: object | None
    table: object | None
    table_caption_p: object | None
    body_p: object | None
    dash_p: object | None
    h1_p: object | None
    h1_center_p: object | None
    h2_p: object | None
    body_style: str
    heading_styles: dict[int, str]
    code_p: object | None
    max_width_emu: int


def _p_text(p_elem) -> str:
    return "".join((t.text or "") for t in p_elem.findall(".//" + qn("w:t")))


def _insert_before_sect(document, elem):
    body = document.element.body
    sect = body.find(qn("w:sectPr"))
    if sect is not None:
        sect.addprevious(elem)
    else:
        body.append(elem)
    return elem


def _style_names(document) -> set[str]:
    names = set()
    for style in document.styles:
        if style.name:
            names.add(style.name)
    return names


WP_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
_MIN_FIGURE_CX = 1_100_000


def _is_toc_line(text: str) -> bool:
    # Leader dots of a TOC line, not ordinary sentences with URLs or versions.
    return "...." in text or bool(re.search(r"\.{5,}", text))


def _p_style_id(p_elem) -> str:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return ""
    st = pPr.find(qn("w:pStyle"))
    return st.get(qn("w:val")) if st is not None else ""


def _has_jc(p_elem, val: str) -> bool:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return False
    jc = pPr.find(qn("w:jc"))
    return jc is not None and jc.get(qn("w:val")) == val


def _drawing_cx(p_elem) -> int:
    ext = p_elem.find(".//" + WP_EXTENT)
    if ext is None or ext.get("cx") is None:
        return 0
    return int(ext.get("cx"))


def collect_prototypes(document, start_elem=None) -> Prototypes:
    body = document.element.body
    children = list(body)
    figure_p = None
    figure_caption_p = None
    table = None
    table_caption_p = None
    body_p = None
    dash_p = None
    h1_p = None
    h1_center_p = None
    h2_p = None
    prev_p = None
    started = start_elem is None
    for child in children:
        if start_elem is not None and child is start_elem:
            started = True
        if not started:
            if child.tag == qn("w:p"):
                prev_p = child
            continue
        if child.tag == qn("w:p"):
            text = _p_text(child).strip()
            if (
                figure_p is None
                and child.find(".//" + qn("w:drawing")) is not None
                and _drawing_cx(child) >= _MIN_FIGURE_CX
            ):
                figure_p = deepcopy(child)
            elif figure_p is not None and figure_caption_p is None and text.startswith("Рисунок"):
                figure_caption_p = deepcopy(child)
            style_id = _p_style_id(child)
            if text.startswith("– ") or text.startswith("— "):
                if dash_p is None:
                    dash_p = deepcopy(child)
            elif style_id in {"1", "Heading1"} and not _is_toc_line(text):
                if _has_jc(child, "center") and h1_center_p is None:
                    h1_center_p = deepcopy(child)
                elif h1_p is None:
                    h1_p = deepcopy(child)
            elif style_id in {"2", "Heading2"} and h2_p is None and not _is_toc_line(text):
                h2_p = deepcopy(child)
            elif (
                body_p is None
                and text
                and child.find(".//" + qn("w:drawing")) is None
                and not text.startswith("Рисунок")
                and not text.startswith("Таблица")
                and style_id not in {"1", "2", "Heading1", "Heading2"}
                and not _is_toc_line(text)
            ):
                body_p = deepcopy(child)
            prev_p = child
        elif child.tag == qn("w:tbl") and table is None:
            table = deepcopy(child)
            if prev_p is not None and _p_text(prev_p).strip().startswith("Таблица"):
                table_caption_p = deepcopy(prev_p)
    if figure_p is not None and figure_caption_p is None:
        figure_caption_p = deepcopy(body_p) if body_p is not None else None
    names = _style_names(document)
    heading_styles = {}
    for level, name in ((1, "Heading 1"), (2, "Heading 2"), (3, "Heading 3")):
        if name in names:
            heading_styles[level] = name
    section = document.sections[0]
    max_width = int(section.page_width - section.left_margin - section.right_margin)
    body_style = "Normal" if "Normal" in names else (document.styles[0].name if document.styles else "Normal")
    return Prototypes(
        figure_p=figure_p,
        figure_caption_p=figure_caption_p,
        table=table,
        table_caption_p=table_caption_p,
        body_p=body_p,
        dash_p=dash_p,
        h1_p=h1_p,
        h1_center_p=h1_center_p,
        h2_p=h2_p,
        body_style=body_style,
        heading_styles=heading_styles,
        code_p=None,
        max_width_emu=max_width,
    )


def _find_body_start(body, title_fields: TitleFields):
    bs = title_fields.body_start
    paras = [c for c in body if c.tag == qn("w:p")]
    if bs.type == "heading_text":
        needle = str(bs.value).strip()
        for p in paras:
            text = _p_text(p).strip()
            if text == needle or text.startswith(needle):
                return p
        raise ComposeError(f"Не найден body_start: {needle}")
    if bs.type == "paragraph_index":
        idx = int(bs.value)
        if idx < 0 or idx >= len(paras):
            raise ComposeError("paragraph_index вне диапазона")
        return paras[idx]
    if bs.type == "section":
        n = int(bs.value)
        # first paragraph of document if n==0; otherwise after n-th sectPr is not
        # easily available in a single body. Treat as heading fallback.
        raise ComposeError("body_start.type=section не поддержан без отдельной секции в теле")
    raise ComposeError(f"Неизвестный body_start.type: {bs.type}")


def _clear_from(body, start_elem) -> None:
    found = False
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        if not found:
            if child is start_elem:
                found = True
                body.remove(child)
            continue
        body.remove(child)
    if not found:
        raise ComposeError("Не удалось очистить тело отчёта")


def _resolve_image(work_dir: Path, raw: str) -> Path:
    path = Path(raw)
    if path.is_file():
        return path
    return (work_dir / raw).resolve()


def _assert_figure_ok(work_dir: Path, raw: str) -> None:
    path = _resolve_image(work_dir, raw)
    if not path.is_file():
        raise ComposeError(f"Нет файла изображения: {raw}")
    posix = Path(raw).as_posix().replace("\\", "/")
    under_shots = posix.startswith("screenshots/") or (work_dir / "screenshots") in path.parents or path.parent == (work_dir / "screenshots")
    if under_shots:
        manifest = load_manifest(work_dir / "screenshots")
        name = path.name
        entry = next((item for item in manifest if item.get("file") == name), None)
        if not entry or not entry.get("ok"):
            raise ComposeError(f"Скриншот не подтверждён: {name}")


def _set_cloned_text(p_elem, text: str) -> None:
    nodes = p_elem.findall(".//" + qn("w:t"))
    if not nodes:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        r.append(t)
        p_elem.append(r)
        return
    nodes[0].text = text
    if text[:1].isspace() or (text and text[-1:].isspace()):
        nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in nodes[1:]:
        node.text = ""


def _first_elem(*elems):
    for el in elems:
        if el is not None:
            return el
    return None


def _set_space_before(p_elem, before: int) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    sp.set(qn("w:before"), str(before))


def _strip_ppr(p_elem, *tags: str) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return
    for tag in tags:
        for el in pPr.findall(qn(f"w:{tag}")):
            pPr.remove(el)


def _ensure_page_break_before(p_elem) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    if pPr.find(qn("w:pageBreakBefore")) is None:
        pPr.append(OxmlElement("w:pageBreakBefore"))


def _ensure_keep_next(p_elem) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    if pPr.find(qn("w:keepNext")) is None:
        pPr.append(OxmlElement("w:keepNext"))


def _ensure_run_bold(p_elem) -> None:
    for r in p_elem.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.insert(0, rPr)
        if rPr.find(qn("w:b")) is None:
            rPr.append(OxmlElement("w:b"))
        if rPr.find(qn("w:bCs")) is None:
            rPr.append(OxmlElement("w:bCs"))


def _set_run_fonts(rPr, name: str) -> None:
    fonts = rPr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rPr.append(fonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        fonts.set(qn(f"w:{attr}"), name)


def _set_run_sz(rPr, sz: str) -> None:
    for tag in ("sz", "szCs"):
        el = rPr.find(qn(f"w:{tag}"))
        if el is None:
            el = OxmlElement(f"w:{tag}")
            rPr.append(el)
        el.set(qn("w:val"), sz)


def _style_heading_run(p_elem, sz: str) -> None:
    _ensure_run_bold(p_elem)
    for r in p_elem.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.insert(0, rPr)
        _set_run_sz(rPr, sz)
        _set_run_fonts(rPr, "Times New Roman")


def _center_p(p_elem) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    _strip_ppr(p_elem, "ind")
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "center")


def _is_unnumbered_heading(text: str) -> bool:
    return not re.match(r"^\d+", text.strip())


def _heading_starts_new_page(text: str) -> bool:
    t = " ".join(text.strip().upper().split())
    if t in {"ЗАКЛЮЧЕНИЕ", "ПРИЛОЖЕНИЯ"}:
        return True
    if t.startswith("СПИСОК ИСПОЛЬЗОВАН") or t.startswith("СПИСОК ИСПОЛЬЗУЕМ"):
        return True
    return False


def _is_listing_caption(text: str) -> bool:
    return bool(re.match(r"^Листинг\b", text.strip()))


def _emit_listing_caption(document, proto: Prototypes, text: str) -> None:
    elem = _add_cloned(
        document,
        _first_elem(proto.table_caption_p, proto.body_p),
        text,
        proto.body_style,
    )
    _style_listing_caption(elem)
    _ensure_keep_next(elem)


def _add_cloned(
    document,
    proto,
    text: str,
    fallback_style: str | None,
    *,
    keep_page_break: bool = False,
    bold: bool = False,
):
    if proto is None:
        elem = _add_paragraph_style(document, text, fallback_style)
        if bold:
            _ensure_run_bold(elem)
        return elem
    new_p = deepcopy(proto)
    _strip_ppr(new_p, "keepNext", "keepLines", "numPr")
    if not keep_page_break:
        _strip_ppr(new_p, "pageBreakBefore")
    else:
        _ensure_page_break_before(new_p)
    _set_cloned_text(new_p, text)
    if bold:
        _ensure_run_bold(new_p)
    return _insert_before_sect(document, new_p)


def _add_paragraph_style(document, text: str, style: str | None):
    kwargs = {}
    if style:
        kwargs["style"] = style
    p = document.add_paragraph(text, **kwargs)
    elem = p._element
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)
    return _insert_before_sect(document, elem)


def _style_code_paragraph(p_elem) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    _strip_ppr(p_elem, "keepNext", "keepLines", "numPr", "pageBreakBefore")
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        pPr.remove(ind)
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "left")
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    for r in p_elem.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.insert(0, rPr)
        fonts = rPr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rPr.append(fonts)
        fonts.set(qn("w:ascii"), "Courier New")
        fonts.set(qn("w:hAnsi"), "Courier New")
        fonts.set(qn("w:cs"), "Courier New")
        fonts.set(qn("w:eastAsia"), "Courier New")
        for tag in ("b", "bCs", "i", "iCs"):
            for el in rPr.findall(qn(f"w:{tag}")):
                rPr.remove(el)
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), "20")
        sz_cs = rPr.find(qn("w:szCs"))
        if sz_cs is None:
            sz_cs = OxmlElement("w:szCs")
            rPr.append(sz_cs)
        sz_cs.set(qn("w:val"), "20")


def _add_code(document, text: str, proto: Prototypes) -> None:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines:
        lines = [""]
    if proto.table is None:
        proto_p = proto.body_p
        for line in lines:
            if proto_p is None:
                elem = _add_paragraph_style(document, line or " ", proto.body_style)
            else:
                new_p = deepcopy(proto_p)
                _set_cloned_text(new_p, line if line else " ")
                elem = _insert_before_sect(document, new_p)
            _style_code_paragraph(elem)
        return
    tbl = deepcopy(proto.table)
    wanted = [[line if line else " "] for line in lines]
    _reshape_table(tbl, wanted)
    _fit_table_width(tbl, _content_dxa(document), 1, col_weights=[1])
    _set_outer_frame(tbl)
    trs = tbl.findall(qn("w:tr"))
    for tr in trs:
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tr.insert(0, trPr)
        for hdr in trPr.findall(qn("w:tblHeader")):
            trPr.remove(hdr)
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        for tc in tr.findall(qn("w:tc")):
            _style_listing_cell(tc, first=(tr is trs[0]), last=(tr is trs[-1]))
    _insert_before_sect(document, tbl)


def _set_cell_text(tc, text: str) -> None:
    for t in tc.findall(".//" + qn("w:t")):
        t.text = ""
    paras = tc.findall(qn("w:p"))
    if not paras:
        p = OxmlElement("w:p")
        tc.append(p)
        paras = [p]
    p = paras[0]
    runs = p.findall(qn("w:r"))
    if not runs:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        r.append(t)
        p.append(r)
        return
    t_nodes = runs[0].findall(qn("w:t"))
    if not t_nodes:
        t = OxmlElement("w:t")
        t.text = text
        runs[0].append(t)
    else:
        t_nodes[0].text = text
    for extra_p in paras[1:]:
        tc.remove(extra_p)


def _content_dxa(document) -> int:
    section = document.sections[0]
    return max(1, int((section.page_width - section.left_margin - section.right_margin) / 635))


def _reshape_table(tbl, wanted: list[list[str]]) -> None:
    trs = tbl.findall(qn("w:tr"))
    while len(trs) < len(wanted) and trs:
        tbl.append(deepcopy(trs[-1]))
        trs = tbl.findall(qn("w:tr"))
    for extra in trs[len(wanted):]:
        tbl.remove(extra)
    trs = tbl.findall(qn("w:tr"))
    for r_idx, row in enumerate(wanted):
        tcs = trs[r_idx].findall(qn("w:tc"))
        while len(tcs) < len(row) and tcs:
            trs[r_idx].append(deepcopy(tcs[-1]))
            tcs = trs[r_idx].findall(qn("w:tc"))
        for extra in list(tcs[len(row):]):
            trs[r_idx].remove(extra)
        tcs = trs[r_idx].findall(qn("w:tc"))
        for c_idx, value in enumerate(row):
            if c_idx < len(tcs):
                _set_cell_text(tcs[c_idx], value)


def _tbl_borders(tbl):
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    return borders


def _set_border(borders, name: str, val: str, sz: str = "4") -> None:
    el = borders.find(qn(f"w:{name}"))
    if el is None:
        el = OxmlElement(f"w:{name}")
        borders.append(el)
    el.set(qn("w:val"), val)
    if val in {"nil", "none"}:
        return
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), "000000")


def _set_outer_frame(tbl) -> None:
    borders = _tbl_borders(tbl)
    for name in ("top", "left", "bottom", "right"):
        _set_border(borders, name, "single")
    _set_border(borders, "insideH", "nil")
    _set_border(borders, "insideV", "nil")


def _nil_bottom_border(tbl) -> None:
    borders = _tbl_borders(tbl)
    _set_border(borders, "bottom", "nil")
    trs = tbl.findall(qn("w:tr"))
    if not trs:
        return
    for tc in trs[-1].findall(qn("w:tc")):
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        tc_borders = tcPr.find(qn("w:tcBorders"))
        if tc_borders is None:
            tc_borders = OxmlElement("w:tcBorders")
            tcPr.append(tc_borders)
        _set_border(tc_borders, "bottom", "nil")


def _style_listing_cell(tc, *, first: bool, last: bool) -> None:
    _set_cell_white(tc)
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tc_borders = tcPr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tcPr.append(tc_borders)
    _set_border(tc_borders, "left", "single")
    _set_border(tc_borders, "right", "single")
    _set_border(tc_borders, "top", "single" if first else "nil")
    _set_border(tc_borders, "bottom", "single" if last else "nil")
    for p in tc.findall(qn("w:p")):
        _style_code_paragraph(p)


def _style_listing_caption(p_elem) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    _strip_ppr(p_elem, "ind", "pageBreakBefore")
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "left")
    for r in p_elem.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.insert(0, rPr)
        for tag in ("b", "bCs"):
            for el in rPr.findall(qn(f"w:{tag}")):
                rPr.remove(el)
        if rPr.find(qn("w:i")) is None:
            rPr.append(OxmlElement("w:i"))
        if rPr.find(qn("w:iCs")) is None:
            rPr.append(OxmlElement("w:iCs"))
        _set_run_sz(rPr, "24")
        _set_run_fonts(rPr, "Times New Roman")


def _set_cell_white(tc) -> None:
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "FFFFFF")


def _normalize_col_widths(widths: list[int], total: int, min_dxa: int) -> list[int]:
    n = len(widths)
    if n == 0 or total <= 0:
        return widths
    min_dxa = max(200, min(min_dxa, total // max(n * 2, 1)))
    widths = [max(min_dxa, int(w)) for w in widths]
    diff = total - sum(widths)
    idx = max(range(n), key=lambda i: widths[i])
    widths[idx] = max(min_dxa, widths[idx] + diff)
    s = sum(widths)
    if s != total and s > 0:
        widths = [max(min_dxa, int(w * total / s)) for w in widths]
        widths[-1] += total - sum(widths)
        if widths[-1] < min_dxa:
            steal = min_dxa - widths[-1]
            widths[-1] = min_dxa
            j = max(range(n - 1), key=lambda i: widths[i]) if n > 1 else 0
            widths[j] = max(min_dxa, widths[j] - steal)
    return [max(1, int(w)) for w in widths]


def _content_col_weights(headers: list[str], rows: list[list[str]], ncols: int) -> list[int]:
    weights: list[int] = []
    for i in range(ncols):
        header_len = len(headers[i]) if i < len(headers) else 1
        longest = header_len * 2
        for row in rows:
            if i < len(row):
                cell = row[i].strip()
                n = min(len(cell), 24) if cell.startswith("http") else len(cell)
                longest = max(longest, n)
        weights.append(max(longest, 2))
    return weights


def _fit_table_width(
    tbl,
    width_dxa: int,
    ncols: int,
    col_weights: list[int] | None = None,
) -> None:
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(width_dxa))
    tblW.set(qn("w:type"), "dxa")
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    old = []
    if grid is not None:
        old = [int(col.get(qn("w:w") or "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w") or 0) for col in grid.findall(qn("w:gridCol"))]
    if col_weights is not None and len(col_weights) == ncols and sum(max(w, 1) for w in col_weights) > 0:
        total_w = sum(max(w, 1) for w in col_weights)
        min_dxa = max(1600 if ncols >= 5 else 1400, width_dxa // max(ncols * 6, 1))
        widths = [max(min_dxa, int(width_dxa * max(w, 1) / total_w)) for w in col_weights]
        cap = max(min_dxa, int(width_dxa * (0.45 if ncols >= 5 else 0.5)))
        widths = [min(w, cap) for w in widths]
    elif len(old) == ncols and sum(old) > 0:
        total_w = sum(old)
        widths = [max(200, int(w * width_dxa / total_w)) for w in old]
        min_dxa = 200
    else:
        base = width_dxa // max(ncols, 1)
        widths = [base] * ncols
        min_dxa = max(200, base)
    widths = _normalize_col_widths(widths, width_dxa, min_dxa)
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl.insert(1, grid)
    for col in list(grid.findall(qn("w:gridCol"))):
        grid.remove(col)
    for w in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(w))
        grid.append(col)
    for tr in tbl.findall(qn("w:tr")):
        tcs = tr.findall(qn("w:tc"))
        for extra in tcs[len(widths):]:
            tr.remove(extra)
        tcs = tr.findall(qn("w:tc"))
        for i, tc in enumerate(tcs):
            if i >= len(widths):
                break
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.insert(0, tcW)
            tcW.set(qn("w:w"), str(widths[i]))
            tcW.set(qn("w:type"), "dxa")
            _set_cell_white(tc)


AFTER_TABLE_TWIPS = "340"


def _strip_fields(p_elem) -> None:
    for el in list(p_elem.findall(".//" + qn("w:fldChar"))):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    for el in list(p_elem.findall(".//" + qn("w:instrText"))):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    for el in list(p_elem.findall(".//" + qn("w:fldSimple"))):
        parent = el.getparent()
        if parent is None:
            continue
        idx = list(parent).index(el)
        for child in list(el):
            parent.insert(idx, child)
            idx += 1
        parent.remove(el)


def _insert_table_caption(document, proto: Prototypes, text: str) -> None:
    if proto.table_caption_p is None:
        return
    cap = deepcopy(proto.table_caption_p)
    _strip_fields(cap)
    _set_cloned_text(cap, text)
    _ensure_keep_next(cap)
    _insert_before_sect(document, cap)


def _finish_data_table(tbl, headers: list[str], rows: list[list[str]], width_dxa: int, hide_bottom: bool) -> None:
    wanted = [headers] + rows if headers else rows
    _reshape_table(tbl, wanted)
    tr0 = tbl.findall(qn("w:tr"))
    ncols = len(tr0[0].findall(qn("w:tc"))) if tr0 else max(
        (len(headers) if headers else 0), max((len(r) for r in rows), default=0), 1
    )
    _fit_table_width(tbl, width_dxa, ncols, col_weights=_content_col_weights(headers, rows, ncols))
    trs = tbl.findall(qn("w:tr"))
    for i, tr in enumerate(trs):
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tr.insert(0, trPr)
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        if i == 0 and headers and trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))
    if hide_bottom:
        _nil_bottom_border(tbl)


def _uniform_grid(headers: list[str], rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    ncols = len(headers) if headers else max((len(r) for r in rows), default=0)
    ncols = max(ncols, 1)
    extra = next((i + 1 for i, r in enumerate(rows) if len(r) > ncols), None)
    if extra is not None:
        raise ComposeError(
            f"В таблице строка {extra} с лишней ячейкой (нужно {ncols} столбцов). "
            "Символ «|» внутри ячейки пишите как \\|"
        )
    hdr = (list(headers) + [""] * ncols)[:ncols]
    data = [(list(r) + [""] * ncols)[:ncols] for r in rows]
    return hdr, data


def _add_table(document, proto: Prototypes, caption: str, n: int, headers: list[str], rows: list[list[str]], template: str) -> None:
    if proto.table is None:
        raise ComposeError("В эталоне нет образца таблицы")
    headers, data = _uniform_grid(headers, rows)
    _insert_table_caption(document, proto, template.format(n=n, caption=caption))
    tbl = deepcopy(proto.table)
    _finish_data_table(tbl, headers, data, _content_dxa(document), hide_bottom=False)
    _insert_before_sect(document, tbl)


def replace_body(
    docx_path: Path,
    report_md: Path,
    title_fields: TitleFields,
    work_dir: Path,
) -> None:
    docx_path = Path(docx_path)
    work_dir = Path(work_dir)
    assert_not_sample(docx_path)
    if not title_fields.confirmed:
        raise ComposeError("title-fields.yaml не подтверждён")
    blocks = parse_report_md(Path(report_md).read_text(encoding="utf-8"))
    for block in blocks:
        if block.kind == "figure":
            _assert_figure_ok(work_dir, block.path)
    document = Document(str(docx_path))
    start = _find_body_start(document.element.body, title_fields)
    proto = collect_prototypes(document, start)
    needs_fig = any(b.kind == "figure" for b in blocks)
    if needs_fig and (proto.figure_p is None or proto.figure_caption_p is None):
        raise ComposeError("В эталоне нет образца рисунка")
    needs_tbl = any(b.kind == "table" for b in blocks)
    if needs_tbl and proto.table is None:
        raise ComposeError("В эталоне нет образца таблицы")
    _clear_from(document.element.body, start)
    toc_h = _add_cloned(
        document,
        _first_elem(proto.h1_center_p, proto.h1_p),
        "СОДЕРЖАНИЕ",
        proto.heading_styles.get(1) or proto.body_style,
        keep_page_break=True,
        bold=True,
    )
    _style_heading_run(toc_h, "36")
    _center_p(toc_h)
    fig_n = 0
    tbl_n = 0
    lst_n = 0
    after_table = False
    first_content_h1 = True
    first_numbered_h1 = True

    def _emit_gap(elem) -> None:
        nonlocal after_table
        if after_table and elem is not None:
            _set_space_before(elem, int(AFTER_TABLE_TWIPS))
            after_table = False

    last_visual = None

    def _note_text() -> None:
        nonlocal last_visual
        last_visual = None

    def _note_object() -> None:
        nonlocal last_visual
        if last_visual == "object":
            raise ComposeError(
                "Между рисунком и рисунком, рисунком и таблицей или таблицей и таблицей нужен текст"
            )
        last_visual = "object"

    for block in blocks:
        if block.kind == "heading" and block.text.strip().upper() in {"СОДЕРЖАНИЕ", "CONTENTS"}:
            continue
        if block.kind == "heading":
            _note_text()
            if block.level >= 2:
                elem = _add_cloned(
                    document,
                    _first_elem(proto.h2_p, proto.h1_p),
                    block.text,
                    proto.heading_styles.get(block.level) or proto.body_style,
                    keep_page_break=_heading_starts_new_page(block.text),
                    bold=True,
                )
                _style_heading_run(elem, "32")
            else:
                unnumbered = _is_unnumbered_heading(block.text)
                h_proto = (
                    _first_elem(proto.h1_center_p, proto.h1_p)
                    if unnumbered
                    else proto.h1_p
                )
                keep_break = (
                    first_content_h1
                    or (not unnumbered)
                    or _heading_starts_new_page(block.text)
                )
                elem = _add_cloned(
                    document,
                    h_proto,
                    block.text,
                    proto.heading_styles.get(1) or proto.body_style,
                    keep_page_break=keep_break,
                    bold=True,
                )
                _style_heading_run(elem, "36")
                if unnumbered:
                    _center_p(elem)
                first_content_h1 = False
                if not unnumbered:
                    first_numbered_h1 = False
            _emit_gap(elem)
        elif block.kind == "dash":
            _note_text()
            elem = _add_cloned(
                document,
                proto.dash_p if proto.dash_p is not None else proto.body_p,
                block.text,
                proto.body_style,
            )
            _emit_gap(elem)
        elif block.kind == "paragraph":
            _note_text()
            if _is_toc_line(block.text):
                continue
            if _is_listing_caption(block.text):
                _emit_listing_caption(document, proto, block.text)
                elem = None
            else:
                elem = _add_cloned(document, proto.body_p, block.text, proto.body_style)
            _emit_gap(elem)
        elif block.kind == "figure":
            _note_object()
            fig_n += 1
            append_cloned_figure(
                document,
                proto.figure_p,
                proto.figure_caption_p,
                _resolve_image(work_dir, block.path),
                block.caption,
                fig_n,
                title_fields.figure_caption,
                proto.max_width_emu,
            )
        elif block.kind == "table":
            _note_object()
            tbl_n += 1
            _add_table(
                document,
                proto,
                block.caption,
                tbl_n,
                block.headers,
                block.rows,
                title_fields.table_caption,
            )
            after_table = True
        elif block.kind == "code":
            _note_text()
            if block.caption:
                if _is_listing_caption(block.caption):
                    cap_text = block.caption
                else:
                    lst_n += 1
                    cap_text = title_fields.listing_caption.format(
                        n=lst_n, caption=block.caption
                    )
                _emit_listing_caption(document, proto, cap_text)
            _add_code(document, block.text, proto)
            after_table = True
    insert_word_toc(document, tab_pos_dxa=_content_dxa(document))
    document.save(str(docx_path))


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--docx", required=True)
    p.add_argument("--md", required=True)
    p.add_argument("--title-fields", required=True)
    p.add_argument("--work-dir", required=True)
    args = p.parse_args()
    try:
        replace_body(
            Path(args.docx),
            Path(args.md),
            load_title_fields(Path(args.title_fields)),
            Path(args.work_dir),
        )
    except ComposeError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
