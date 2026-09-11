from pathlib import Path
import hashlib
import pytest
from docx import Document
from tests.fixtures_util import make_sample_docx
from tools.config import BodyStart, Student, TitleField, TitleFields, WorkMeta
from tools.patch_title import TitleFieldNotFound, TitleNotConfirmed, patch_title

TF = TitleFields(
    fields={
        "student_name": TitleField("Петров Пётр Петрович", "student.full_name"),
        "group": TitleField("ИКБО-00-00", "student.group"),
        "topic": TitleField("Лабораторная работа №1. Пример", "work.title"),
    },
    body_start=BodyStart("heading_text", "Цель работы"),
    figure_caption="Рисунок {n} – {caption}",
    table_caption="Таблица {n} – {caption}",
    confirmed=True,
)
STUDENT = Student("Иванов Иван Иванович", "Иванов И.И.", "ИКБО-21-23")
WORK = WorkMeta("Лабораторная работа №3. Процессы", 3, "os")


def test_patch_title_replaces_fields(tmp_path: Path):
    docx = tmp_path / "report.docx"
    make_sample_docx(docx)
    patch_title(docx, TF, STUDENT, WORK)
    texts = [p.text for p in Document(str(docx)).paragraphs]
    joined = "\n".join(texts)
    assert "Иванов Иван Иванович" in joined
    assert "Петров Пётр Петрович" not in joined
    assert "ИКБО-21-23" in joined
    assert "Лабораторная работа №3. Процессы" in joined
    assert "Цель работы" in joined


def test_patch_title_missing_find_leaves_file_untouched(tmp_path: Path):
    docx = tmp_path / "report.docx"
    make_sample_docx(docx)
    before = hashlib.sha256(docx.read_bytes()).hexdigest()
    bad = TitleFields(
        fields={"x": TitleField("ТАКОЙ СТРОКИ НЕТ", "student.full_name")},
        body_start=TF.body_start,
        figure_caption=TF.figure_caption,
        table_caption=TF.table_caption,
        confirmed=True,
    )
    with pytest.raises(TitleFieldNotFound):
        patch_title(docx, bad, STUDENT, WORK)
    assert hashlib.sha256(docx.read_bytes()).hexdigest() == before


def test_patch_title_refuses_unconfirmed(tmp_path: Path):
    docx = tmp_path / "report.docx"
    make_sample_docx(docx)
    un = TitleFields(TF.fields, TF.body_start, TF.figure_caption, TF.table_caption, False)
    with pytest.raises(TitleNotConfirmed, match="не подтверждён"):
        patch_title(docx, un, STUDENT, WORK)
