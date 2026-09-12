from __future__ import annotations

from copy import deepcopy

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips


def _p_text(p_elem) -> str:
    return "".join((t.text or "") for t in p_elem.findall(".//" + qn("w:t"))).strip()


def _p_style(p_elem) -> str:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return ""
    st = pPr.find(qn("w:pStyle"))
    return (st.get(qn("w:val")) or "") if st is not None else ""


def _ensure_style(document, name: str, style_id: str, based_on: str = "Normal"):
    try:
        return document.styles[name]
    except KeyError:
        pass
    style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    try:
        style.base_style = document.styles[based_on]
    except KeyError:
        pass
    el = style._element
    el.set(qn("w:styleId"), style_id)
    return style


def ensure_toc_styles(document, tab_pos_dxa: int) -> None:
    toc1 = _ensure_style(document, "TOC 1", "TOC1")
    toc2 = _ensure_style(document, "TOC 2", "TOC2")
    for style, left in ((toc1, 0), (toc2, 283)):
        pf = style.paragraph_format
        pf.line_spacing = 1.5
        pf.first_line_indent = Twips(0)
        pf.left_indent = Twips(left)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pPr = style._element.get_or_add_pPr()
        tabs = pPr.find(qn("w:tabs"))
        if tabs is None:
            tabs = OxmlElement("w:tabs")
            pPr.append(tabs)
        for old in list(tabs):
            tabs.remove(old)
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:leader"), "dot")
        tab.set(qn("w:pos"), str(tab_pos_dxa))
        tabs.append(tab)
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.bold = False


def _bookmark_heading(p_elem, name: str, bid: str) -> None:
    for old in list(p_elem.findall(qn("w:bookmarkStart"))):
        if old.get(qn("w:name")) == name:
            return
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), bid)
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bid)
    p_elem.insert(0, start)
    p_elem.append(end)


def _toc_entry_paragraph(text: str, level: int, bookmark: str, page_hint: str = "1"):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), "TOC1" if level <= 1 else "TOC2")
    pPr.append(pStyle)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), "9356")
    tabs.append(tab)
    pPr.append(tabs)
    p.append(pPr)
    hyper = OxmlElement("w:hyperlink")
    hyper.set(qn("w:anchor"), bookmark)
    hyper.set(qn("w:history"), "1")

    def run_text(value: str, extra_rpr=None):
        r = OxmlElement("w:r")
        if extra_rpr is not None:
            r.append(deepcopy(extra_rpr))
        t = OxmlElement("w:t")
        t.text = value
        r.append(t)
        return r

    hyper.append(run_text(text))
    r_tab = OxmlElement("w:r")
    r_tab.append(OxmlElement("w:tab"))
    hyper.append(r_tab)

    def fld(kind: str):
        r = OxmlElement("w:r")
        ch = OxmlElement("w:fldChar")
        ch.set(qn("w:fldCharType"), kind)
        r.append(ch)
        return r

    hyper.append(fld("begin"))
    r_instr = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = f" PAGEREF {bookmark} \\h "
    r_instr.append(instr)
    hyper.append(r_instr)
    hyper.append(fld("separate"))
    hyper.append(run_text(page_hint))
    hyper.append(fld("end"))
    p.append(hyper)
    return p


def _field_run(kind: str):
    r = OxmlElement("w:r")
    ch = OxmlElement("w:fldChar")
    ch.set(qn("w:fldCharType"), kind)
    if kind == "begin":
        ch.set(qn("w:dirty"), "true")
    r.append(ch)
    return r


def insert_word_toc(document, tab_pos_dxa: int = 9356) -> None:
    """Insert a Word-native TOC (SDT + TOC field) after the СОДЕРЖАНИЕ heading."""
    ensure_toc_styles(document, tab_pos_dxa)
    body = document.element.body
    contents_p = None
    for child in list(body):
        if child.tag != qn("w:p"):
            continue
        if _p_text(child).upper() == "СОДЕРЖАНИЕ":
            contents_p = child
            break
    if contents_p is None:
        return
    nxt = contents_p.getnext()
    while nxt is not None and (
        nxt.tag == qn("w:sdt")
        or (
            nxt.tag == qn("w:p")
            and (
                nxt.find(".//" + qn("w:instrText")) is not None
                or _p_style(nxt).startswith("TOC")
                or _p_text(nxt) in {"Оглавление", ""}
            )
        )
    ):
        following = nxt.getnext()
        body.remove(nxt)
        nxt = following

    headings = []
    started = False
    bid = 1000
    for child in list(body):
        if child is contents_p:
            started = True
            continue
        if not started or child.tag != qn("w:p"):
            continue
        style = _p_style(child)
        text = _p_text(child)
        if not text or text.upper() == "СОДЕРЖАНИЕ":
            continue
        if style in {"1", "Heading1"}:
            name = f"_Toc{bid}"
            _bookmark_heading(child, name, str(bid))
            headings.append((text, 1, name))
            bid += 1
        elif style in {"2", "Heading2"}:
            name = f"_Toc{bid}"
            _bookmark_heading(child, name, str(bid))
            headings.append((text, 2, name))
            bid += 1

    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    sid = OxmlElement("w:id")
    sid.set(qn("w:val"), "351654032")
    sdtPr.append(sid)
    doc_part = OxmlElement("w:docPartObj")
    gallery = OxmlElement("w:docPartGallery")
    gallery.set(qn("w:val"), "Table of Contents")
    unique = OxmlElement("w:docPartUnique")
    doc_part.append(gallery)
    doc_part.append(unique)
    sdtPr.append(doc_part)
    sdt.append(sdtPr)
    content = OxmlElement("w:sdtContent")

    begin_p = OxmlElement("w:p")
    begin_p.append(_field_run("begin"))
    r_instr = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = ' TOC \\o "1-2" \\h \\z \\u '
    r_instr.append(instr)
    begin_p.append(r_instr)
    begin_p.append(_field_run("separate"))
    content.append(begin_p)

    for text, level, name in headings:
        content.append(_toc_entry_paragraph(text, level, name))

    end_p = OxmlElement("w:p")
    end_p.append(_field_run("end"))
    content.append(end_p)
    sdt.append(content)
    contents_p.addnext(sdt)
