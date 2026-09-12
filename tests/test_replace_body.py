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


def _ppr_attr(paragraph, child, attr):
    from docx.oxml.ns import qn
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    el = pPr.find(qn(f"w:{child}"))
    if el is None:
        return None
    return el.get(qn(f"w:{attr}"))


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
        "| 1 | 2 | 3 | 4 | 5 |\n",
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
