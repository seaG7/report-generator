from __future__ import annotations

import argparse
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


class ComposeError(Exception):
    pass


@dataclass
class Prototypes:
    figure_p: object | None
    figure_caption_p: object | None
    table: object | None
    table_caption_p: object | None
    body_style: str
    heading_styles: dict[int, str]
    code_p: object | None
    max_width_emu: int


def _p_text(p_elem) -> str:
    return "".join((t.text or "") for t in p_elem.findall(".//" + qn("w:t")))


def _insert_before_sect(document, elem) -> None:
    body = document.element.body
    sect = body.find(qn("w:sectPr"))
    if sect is not None:
        sect.addprevious(elem)
    else:
        body.append(elem)


def _style_names(document) -> set[str]:
    names = set()
    for style in document.styles:
        if style.name:
            names.add(style.name)
    return names


def collect_prototypes(document) -> Prototypes:
    body = document.element.body
    children = list(body)
    figure_p = None
    figure_caption_p = None
    table = None
    table_caption_p = None
    prev_p = None
    for child in children:
        if child.tag == qn("w:p"):
            if figure_p is None and child.find(".//" + qn("w:drawing")) is not None:
                figure_p = deepcopy(child)
            elif figure_p is not None and figure_caption_p is None:
                figure_caption_p = deepcopy(child)
            prev_p = child
        elif child.tag == qn("w:tbl") and table is None:
            table = deepcopy(child)
            if prev_p is not None and _p_text(prev_p).strip().startswith("Таблица"):
                table_caption_p = deepcopy(prev_p)
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


def _add_paragraph(document, text: str, style: str | None) -> None:
    kwargs = {}
    if style:
        kwargs["style"] = style
    p = document.add_paragraph(text, **kwargs)
    elem = p._element
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)
    _insert_before_sect(document, elem)


def _add_code(document, text: str, proto: Prototypes) -> None:
    p = document.add_paragraph(text, style=proto.body_style)
    if p.runs:
        rPr = p.runs[0]._element.get_or_add_rPr()
        fonts = rPr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rPr.append(fonts)
        fonts.set(qn("w:ascii"), "Consolas")
        fonts.set(qn("w:hAnsi"), "Consolas")
    elem = p._element
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)
    _insert_before_sect(document, elem)


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


def _add_table(document, proto: Prototypes, caption: str, n: int, headers: list[str], rows: list[list[str]], template: str) -> None:
    if proto.table is None:
        raise ComposeError("В эталоне нет образца таблицы")
    if proto.table_caption_p is not None:
        cap = deepcopy(proto.table_caption_p)
        nodes = cap.findall(".//" + qn("w:t"))
        text = template.format(n=n, caption=caption)
        if nodes:
            nodes[0].text = text
            for node in nodes[1:]:
                node.text = ""
        _insert_before_sect(document, cap)
    tbl = deepcopy(proto.table)
    trs = tbl.findall(qn("w:tr"))
    wanted = [headers] + rows if headers else rows
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
        for c_idx, value in enumerate(row):
            if c_idx < len(tcs):
                _set_cell_text(tcs[c_idx], value)
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
    proto = collect_prototypes(document)
    needs_fig = any(b.kind == "figure" for b in blocks)
    if needs_fig and (proto.figure_p is None or proto.figure_caption_p is None):
        raise ComposeError("В эталоне нет образца рисунка")
    needs_tbl = any(b.kind == "table" for b in blocks)
    if needs_tbl and proto.table is None:
        raise ComposeError("В эталоне нет образца таблицы")
    start = _find_body_start(document.element.body, title_fields)
    _clear_from(document.element.body, start)
    fig_n = 0
    tbl_n = 0
    for block in blocks:
        if block.kind == "heading":
            style = proto.heading_styles.get(block.level)
            _add_paragraph(document, block.text, style or proto.body_style)
        elif block.kind == "paragraph":
            _add_paragraph(document, block.text, proto.body_style)
        elif block.kind == "figure":
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
        elif block.kind == "code":
            _add_code(document, block.text, proto)
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
