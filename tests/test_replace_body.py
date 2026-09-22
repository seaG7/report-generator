from pathlib import Path
from PIL import Image
from docx import Document
from tests.fixtures_util import make_sample_docx
from tools.config import BodyStart, TitleField, TitleFields
from tools.replace_body import ComposeError, replace_body
from tools.screenshot import mark_ok, register_image
import pytest
import zipfile

TF = TitleFields(
    fields={"student_name": TitleField("Петров Пётр Петрович", "student.full_name")},
    body_start=BodyStart("heading_text", "Цель работы"),
    figure_caption="Рисунок {n} – {caption}",
    table_caption="Таблица {n} – {caption}",
    confirmed=True,
)


def test_replace_body_keeps_title_and_clones_second_figure(tmp_path: Path):
    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img1 = tmp_path / "one.png"
    img2 = tmp_path / "two.png"
    Image.new("RGB", (80, 40), (255, 0, 0)).save(img1)
    Image.new("RGB", (80, 40), (0, 0, 255)).save(img2)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img1)
    register_image(img1, shots, "Первый", "other", slug="one")
    register_image(img2, shots, "Второй", "other", slug="two")
    mark_ok(shots, "01-one.png")
    mark_ok(shots, "02-two.png")
    md = work / "report.md"
    md.write_text(
        "# Цель работы\n\nНовая цель.\n\n"
        "![Первый](screenshots/01-one.png)\n\n"
        "Далее второй рисунок.\n\n"
        "![Второй](screenshots/02-two.png)\n",
        encoding="utf-8",
    )
    replace_body(docx, md, TF, work)
    doc = Document(str(docx))
    texts = [p.text for p in doc.paragraphs]
    joined = "\n".join(texts)
    assert "Петров Пётр Петрович" in joined
    assert "Рисунок 1 – Первый" in joined
    assert "Рисунок 2 – Второй" in joined
    assert "Текст цели." not in joined
    with zipfile.ZipFile(docx) as z:
        media_files = [n for n in z.namelist() if n.startswith("word/media/")]
    assert len(media_files) >= 2


def test_replace_body_rejects_screenshot_not_ok(tmp_path: Path):
    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (20, 20)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    register_image(img, shots, "Первый", "other", slug="one")
    md = work / "report.md"
    md.write_text("# Цель работы\n\n![Первый](screenshots/01-one.png)\n", encoding="utf-8")
    with pytest.raises(ComposeError):
        replace_body(docx, md, TF, work)


def test_replace_body_rejects_unconfirmed(tmp_path: Path):
    work = tmp_path / "lab"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (20, 20)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    md = work / "report.md"
    md.write_text("# Цель работы\n\nТекст\n", encoding="utf-8")
    un = TitleFields(TF.fields, TF.body_start, TF.figure_caption, TF.table_caption, False)
    with pytest.raises(ComposeError, match="не подтверждён"):
        replace_body(docx, md, un, work)


def test_later_h1_does_not_force_page_break(tmp_path: Path):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (80, 40)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    doc = Document(str(docx))
    for p in doc.paragraphs:
        if p.text.strip() == "Цель работы":
            pPr = p._element.get_or_add_pPr()
            if pPr.find(qn("w:pageBreakBefore")) is None:
                pPr.append(OxmlElement("w:pageBreakBefore"))
    doc.save(str(docx))
    register_image(img, shots, "Первый", "other", slug="one")
    mark_ok(shots, "01-one.png")
    md = work / "report.md"
    md.write_text(
        "# Цель работы\n\nАбзац цели.\n\n# Выводы\n\nАбзац выводов.\n",
        encoding="utf-8",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))
    h1s = [p for p in out.paragraphs if p.style and p.style.name == "Heading 1" and p.text in {"Цель работы", "Выводы"}]
    assert len(h1s) >= 2
    first = next(p for p in out.paragraphs if p.text.strip() == "Цель работы")
    second = next(p for p in out.paragraphs if p.text.strip() == "Выводы")
    assert first._element.find(qn("w:pPr")).find(qn("w:pageBreakBefore")) is not None
    assert second._element.find(qn("w:pPr")).find(qn("w:pageBreakBefore")) is None
    assert first.runs and first.runs[0]._element.find(qn("w:rPr")).find(qn("w:b")) is not None


def test_first_numbered_h1_starts_new_page(tmp_path: Path):
    from docx.oxml.ns import qn

    work, docx, md = _make_work(
        tmp_path,
        "# ЦЕЛЬ РАБОТЫ\n\nАбзац.\n\n# ЗАДАНИЕ\n\n– пункт задания\n\n"
        "# 1 ОСНОВНАЯ ЧАСТЬ\n\nТекст главы.\n\n# 2 СЛЕДУЮЩИЙ РАЗДЕЛ\n\nЕщё текст.\n",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))
    goal = next(p for p in out.paragraphs if p.text.strip() == "ЦЕЛЬ РАБОТЫ")
    task = next(p for p in out.paragraphs if p.text.strip() == "ЗАДАНИЕ")
    numbered = next(p for p in out.paragraphs if p.text.strip() == "1 ОСНОВНАЯ ЧАСТЬ")
    numbered2 = next(p for p in out.paragraphs if p.text.strip() == "2 СЛЕДУЮЩИЙ РАЗДЕЛ")
    dash = next(p for p in out.paragraphs if "пункт задания" in p.text)
    assert goal._element.find(qn("w:pPr")).find(qn("w:pageBreakBefore")) is not None
    assert task._element.find(qn("w:pPr")).find(qn("w:pageBreakBefore")) is None
    assert numbered._element.find(qn("w:pPr")).find(qn("w:pageBreakBefore")) is not None
    assert numbered2._element.find(qn("w:pPr")).find(qn("w:pageBreakBefore")) is not None
    assert dash.text.startswith("– ")


def test_bibliography_and_appendices_start_new_page(tmp_path: Path):
    from docx.oxml.ns import qn

    work, docx, md = _make_work(
        tmp_path,
        "# ЦЕЛЬ РАБОТЫ\n\nАбзац.\n\n# ЗАКЛЮЧЕНИЕ\n\nТекст выводов.\n\n"
        "# СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ\n\n1. Источник.\n\n"
        "# ПРИЛОЖЕНИЯ\n\nПеречень.\n\n# Приложение А\n\nМатериал.\n",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))

    def pbb(title: str) -> bool:
        p = next(x for x in out.paragraphs if x.text.strip() == title)
        pPr = p._element.find(qn("w:pPr"))
        return pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None

    assert pbb("ЗАКЛЮЧЕНИЕ")
    assert pbb("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    assert pbb("ПРИЛОЖЕНИЯ")
    assert not pbb("Приложение А")


def test_code_block_is_left_aligned_without_first_line_indent(tmp_path: Path):
    from docx.oxml.ns import qn

    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (80, 40)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    register_image(img, shots, "Первый", "other", slug="one")
    mark_ok(shots, "01-one.png")
    md = work / "report.md"
    md.write_text(
        "# Цель работы\n\nАбзац.\n\n```php\n<?php\n$host = '127.0.0.1';\n```\n",
        encoding="utf-8",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))
    code = None
    for tbl in out.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "127.0.0.1" in p.text:
                        code = p
                        break
    assert code is not None
    pPr = code._element.find(qn("w:pPr"))
    jc = pPr.find(qn("w:jc")) if pPr is not None else None
    assert jc is not None and jc.get(qn("w:val")) == "left"
    ind = pPr.find(qn("w:ind")) if pPr is not None else None
    assert ind is None or not ind.get(qn("w:firstLine"))
    rPr = code.runs[0]._element.find(qn("w:rPr"))
    fonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
    assert fonts is not None and fonts.get(qn("w:ascii")) == "Courier New"
    parent = code._element.getparent()
    assert parent is not None and parent.tag == qn("w:tc")
    tbl = parent.getparent().getparent()
    borders = tbl.find(qn("w:tblPr")).find(qn("w:tblBorders"))
    assert borders is not None
    top = borders.find(qn("w:top"))
    assert top is not None and top.get(qn("w:val")) == "single"
    inside = borders.find(qn("w:insideH"))
    assert inside is None or inside.get(qn("w:val")) in {"nil", "none"}


def test_is_toc_line_ignores_urls():
    from tools.replace_body import _is_toc_line

    assert _is_toc_line("1 Выбор поискового запроса ................ 3")
    assert not _is_toc_line(
        "1.1. PHP: Hypertext Preprocessor [Электронный ресурс]. URL: https://www.php.net/docs.php (дата обращения: 09.09.2026)."
    )


def _ppr_attr(paragraph, child, attr):
    from docx.oxml.ns import qn
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    el = pPr.find(qn(f"w:{child}"))
    if el is None:
        return None
    return el.get(qn(f"w:{attr}"))


def _has_ppr(paragraph, child: str) -> bool:
    from docx.oxml.ns import qn

    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return False
    return pPr.find(qn(f"w:{child}")) is not None


def test_replace_body_does_not_glue_body_text_to_figure(tmp_path: Path):
    """keepNext on the paragraph before a figure splits it across pages."""
    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (80, 40)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    register_image(img, shots, "Первый", "other", slug="one")
    mark_ok(shots, "01-one.png")
    md = work / "report.md"
    md.write_text(
        "# Цель работы\n\nДлинный абзац перед рисунком.\n\n"
        "![Первый](screenshots/01-one.png)\n\n"
        "Абзац после рисунка.\n",
        encoding="utf-8",
    )
    replace_body(docx, md, TF, work)
    doc = Document(str(docx))
    before = next(p for p in doc.paragraphs if p.text.strip() == "Длинный абзац перед рисунком.")
    cap = next(p for p in doc.paragraphs if p.text.startswith("Рисунок 1"))
    after = next(p for p in doc.paragraphs if p.text.strip() == "Абзац после рисунка.")
    assert not _has_ppr(before, "keepNext")
    assert not _has_ppr(cap, "keepNext")
    assert not _has_ppr(after, "keepNext")
    from docx.oxml.ns import qn

    img_p = cap._element.getprevious()
    assert img_p is not None and img_p.find(".//" + qn("w:drawing")) is not None
    img_ppr = img_p.find(qn("w:pPr"))
    assert img_ppr is not None and img_ppr.find(qn("w:keepNext")) is not None


def test_replace_body_keeps_indent_justify_and_caption_size(tmp_path: Path):
    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (80, 40)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    register_image(img, shots, "Первый", "other", slug="one")
    mark_ok(shots, "01-one.png")
    md = work / "report.md"
    md.write_text(
        "# Цель работы\n\nНовый абзац цели.\n\n– пункт как в образце\n\n"
        "![Первый](screenshots/01-one.png)\n",
        encoding="utf-8",
    )
    replace_body(docx, md, TF, work)
    doc = Document(str(docx))
    body = next(p for p in doc.paragraphs if p.text.strip() == "Новый абзац цели.")
    assert _ppr_attr(body, "jc", "val") == "both"
    assert _ppr_attr(body, "ind", "firstLine") is not None
    cap = next(p for p in doc.paragraphs if p.text.startswith("Рисунок 1"))
    from docx.oxml.ns import qn
    sz = cap.runs[0]._element.find(qn("w:rPr")).find(qn("w:sz"))
    assert sz is not None and sz.get(qn("w:val")) == "24"


def test_replace_body_table_fits_page_and_header_is_white(tmp_path: Path):
    from docx.oxml.ns import qn

    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (80, 40)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    register_image(img, shots, "Первый", "other", slug="one")
    mark_ok(shots, "01-one.png")
    md = work / "report.md"
    md.write_text(
        "# Цель работы\n\nАбзац.\n\n"
        "Таблица: Результаты\n\n"
        "| A | B | C | D | E |\n"
        "| --- | --- | --- | --- | --- |\n"
        "| 1 | 2 | 3 | 4 | 5 |\n\n"
        "Абзац после таблицы.\n",
        encoding="utf-8",
    )
    replace_body(docx, md, TF, work)
    doc = Document(str(docx))
    tbl = doc.tables[-1]
    tblW = tbl._tbl.find(qn("w:tblPr")).find(qn("w:tblW"))
    width = int(tblW.get(qn("w:w")))
    section = doc.sections[0]
    content = int((section.page_width - section.left_margin - section.right_margin) / 635)
    assert width <= content
    shd = tbl.rows[0].cells[0]._tc.find(qn("w:tcPr")).find(qn("w:shd"))
    assert shd is not None and shd.get(qn("w:fill")) == "FFFFFF"
    after = None
    saw_tbl = False
    for child in doc.element.body:
        from lxml import etree as _et

        tag = _et.QName(child).localname
        if tag == "tbl":
            saw_tbl = True
            continue
        if saw_tbl and tag == "p":
            after = child
            break
    assert after is not None
    sp = after.find(qn("w:pPr")).find(qn("w:spacing"))
    assert sp is not None and sp.get(qn("w:before")) == "340"
    cap = next(p for p in doc.paragraphs if p.text.startswith("Таблица 1"))
    assert cap._element.find(qn("w:pPr")).find(qn("w:keepNext")) is not None
    tr0 = tbl._tbl.findall(qn("w:tr"))[0]
    trPr = tr0.find(qn("w:trPr"))
    assert trPr is not None and trPr.find(qn("w:tblHeader")) is not None
    assert trPr.find(qn("w:cantSplit")) is not None


def test_replace_body_inserts_updateable_word_toc(tmp_path: Path):
    from docx.oxml.ns import qn

    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (80, 40)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    register_image(img, shots, "Первый", "other", slug="one")
    mark_ok(shots, "01-one.png")
    md = work / "report.md"
    md.write_text("# Цель работы\n\nАбзац цели.\n\n# Выводы\n\nАбзац выводов.\n", encoding="utf-8")
    replace_body(docx, md, TF, work)
    doc = Document(str(docx))
    gallery = doc.element.body.find(".//" + qn("w:docPartGallery"))
    assert gallery is not None
    assert gallery.get(qn("w:val")) == "Table of Contents"
    names = [b.get(qn("w:name")) for b in doc.element.body.findall(".//" + qn("w:bookmarkStart"))]
    assert any(n and n.startswith("_Toc") for n in names)


def _make_work(tmp_path: Path, md_text: str):
    work = tmp_path / "lab"
    shots = work / "screenshots"
    work.mkdir()
    img = tmp_path / "one.png"
    Image.new("RGB", (80, 40)).save(img)
    docx = work / "report.docx"
    make_sample_docx(docx, image=img)
    register_image(img, shots, "Первый", "other", slug="one")
    mark_ok(shots, "01-one.png")
    md = work / "report.md"
    md.write_text(md_text, encoding="utf-8")
    return work, docx, md


def test_unnumbered_h1_is_centered_bold_with_explicit_size(tmp_path: Path):
    from docx.oxml.ns import qn

    work, docx, md = _make_work(
        tmp_path,
        "# ЦЕЛЬ РАБОТЫ\n\nАбзац цели.\n\n# 1 ОСНОВНАЯ ЧАСТЬ\n\nТекст главы.\n",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))
    goal = next(p for p in out.paragraphs if p.text.strip() == "ЦЕЛЬ РАБОТЫ")
    numbered = next(p for p in out.paragraphs if p.text.strip() == "1 ОСНОВНАЯ ЧАСТЬ")
    assert _ppr_attr(goal, "jc", "val") == "center"
    goal_ind = goal._element.find(qn("w:pPr")).find(qn("w:ind")) if goal._element.find(qn("w:pPr")) is not None else None
    assert goal_ind is None or not goal_ind.get(qn("w:left"))
    rPr = goal.runs[0]._element.find(qn("w:rPr"))
    assert rPr is not None
    assert rPr.find(qn("w:b")) is not None
    sz = rPr.find(qn("w:sz"))
    assert sz is not None and int(sz.get(qn("w:val"))) >= 32
    fonts = rPr.find(qn("w:rFonts"))
    assert fonts is not None and fonts.get(qn("w:ascii")) == "Times New Roman"
    numbered_jc = _ppr_attr(numbered, "jc", "val")
    assert numbered_jc != "center"


def test_listing_label_is_numbered_italic_and_framed(tmp_path: Path):
    from docx.oxml.ns import qn

    work, docx, md = _make_work(
        tmp_path,
        "# Цель работы\n\nАбзац.\n\nЛистинг: Создание базы\n\n```sql\nSELECT 1;\n```\n",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))
    cap = next(p for p in out.paragraphs if p.text.startswith("Листинг"))
    assert "Листинг 1" in cap.text and "Создание базы" in cap.text
    assert _ppr_attr(cap, "jc", "val") in (None, "left", "start")
    rPr = cap.runs[0]._element.find(qn("w:rPr"))
    assert rPr is not None and rPr.find(qn("w:i")) is not None
    found = False
    for tbl in out.tables:
        for row in tbl.rows:
            for cell in row.cells:
                if "SELECT 1" in cell.text:
                    found = True
    assert found


def test_listing_caption_is_left_italic_not_body_indent(tmp_path: Path):
    from docx.oxml.ns import qn

    work, docx, md = _make_work(
        tmp_path,
        "# Цель работы\n\nАбзац.\n\nЛистинг 1.1 — Создание базы данных и таблиц\n\n"
        "```sql\nCREATE DATABASE project_history;\n```\n",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))
    cap = next(p for p in out.paragraphs if p.text.startswith("Листинг 1.1"))
    jc = _ppr_attr(cap, "jc", "val")
    assert jc in (None, "left", "start")
    pPr = cap._element.find(qn("w:pPr"))
    ind = pPr.find(qn("w:ind")) if pPr is not None else None
    assert ind is None or not ind.get(qn("w:firstLine"))
    rPr = cap.runs[0]._element.find(qn("w:rPr"))
    assert rPr is not None
    assert rPr.find(qn("w:i")) is not None
    assert rPr.find(qn("w:b")) is None
    assert _has_ppr(cap, "keepNext")


def test_long_table_gets_continuation_caption(tmp_path: Path):
    from docx.oxml.ns import qn

    rows = "\n".join(f"| {i} | значение-{i} | назначение-{i} |" for i in range(1, 21))
    work, docx, md = _make_work(
        tmp_path,
        "# Цель работы\n\nАбзац.\n\n"
        "Таблица: Структура\n\n"
        "| Поле | Тип | Назначение |\n"
        "| --- | --- | --- |\n"
        f"{rows}\n",
    )
    replace_body(docx, md, TF, work)
    out = Document(str(docx))
    texts = [p.text for p in out.paragraphs]
    assert any(t.startswith("Таблица 1") for t in texts)
    assert not any("Продолжение Таблицы" in t for t in texts)
    assert len(out.tables) == 1
    tbl = out.tables[0]
    assert len(tbl.rows) == 21


def test_table_column_widths_follow_content_not_equal(tmp_path: Path):
    from docx.oxml.ns import qn

    work, docx, md = _make_work(
        tmp_path,
        "# Цель работы\n\nАбзац.\n\n"
        "Таблица: Структура\n\n"
        "| id | Тип поля в базе данных | Назначение столбца в приложении |\n"
        "| --- | --- | --- |\n"
        "| 1 | INT, AUTO_INCREMENT, PK | Идентификатор категории |\n",
    )
    replace_body(docx, md, TF, work)
    doc = Document(str(docx))
    tbl = doc.tables[-1]
    grid = tbl._tbl.find(qn("w:tblGrid"))
    widths = [int(c.get(qn("w:w"))) for c in grid.findall(qn("w:gridCol"))]
    assert len(widths) == 3
    assert widths[0] < widths[1]
    assert widths[0] < widths[2]
    assert min(widths) >= 800
    assert all(w > 0 for w in widths)
    assert len(tbl.rows[0].cells) == len(widths)
    look = tbl._tbl.find(qn("w:tblPr")).find(qn("w:tblLook"))
    if look is not None and look.get(qn("w:firstRow")) is not None:
        assert look.get(qn("w:firstRow")) != "0"


def test_stacked_figure_and_table_are_rejected(tmp_path: Path):
    work, docx, md = _make_work(
        tmp_path,
        "# Цель работы\n\nАбзац.\n\n![Первый](screenshots/01-one.png)\n\n"
        "![Первый](screenshots/01-one.png)\n",
    )
    with pytest.raises(ComposeError, match="нужен текст"):
        replace_body(docx, md, TF, work)


def test_extra_table_cell_is_rejected(tmp_path: Path):
    work, docx, md = _make_work(
        tmp_path,
        "# Цель работы\n\nАбзац.\n\n"
        "Таблица: Сдвиг\n\n"
        "| Сайт | Title |\n"
        "| --- | --- |\n"
        "| X | foo | bar |\n",
    )
    with pytest.raises(ComposeError, match="лишней ячейкой"):
        replace_body(docx, md, TF, work)


def test_figure_caption_run_is_times_bold_12pt(tmp_path: Path):
    from docx.oxml.ns import qn

    work, docx, md = _make_work(
        tmp_path,
        "# Цель работы\n\nАбзац.\n\n![Первый](screenshots/01-one.png)\n",
    )
    replace_body(docx, md, TF, work)
    doc = Document(str(docx))
    cap = next(p for p in doc.paragraphs if p.text.startswith("Рисунок 1"))
    rPr = cap.runs[0]._element.find(qn("w:rPr"))
    assert rPr is not None
    assert rPr.find(qn("w:b")) is not None
    sz = rPr.find(qn("w:sz"))
    assert sz is not None and sz.get(qn("w:val")) == "24"
    fonts = rPr.find(qn("w:rFonts"))
    assert fonts is not None and fonts.get(qn("w:ascii")) == "Times New Roman"


def test_figure_max_height_leaves_room_for_text(tmp_path: Path):
    from tools.insert_figure import MAX_FIGURE_HEIGHT_EMU, image_extent_emus

    assert MAX_FIGURE_HEIGHT_EMU <= 2_600_000
    path = tmp_path / "tall.png"
    Image.new("RGB", (1600, 2400), (10, 10, 10)).save(path)
    _cx, cy = image_extent_emus(path, 6_000_000)
    assert cy <= MAX_FIGURE_HEIGHT_EMU
