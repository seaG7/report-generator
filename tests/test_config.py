from pathlib import Path
import pytest
import yaml
from tools.config import load_student, load_title_fields, load_work, resolve_to, Student, WorkMeta


def test_load_student_example():
    from tools.paths import repo_root
    s = load_student(repo_root() / "student.yaml.example")
    assert s.full_name == "Иванов Иван Иванович"
    assert s.group == "ИКБО-21-23"


def test_resolve_to_student_and_work():
    student = Student("А А А", "А.А.", "Г1")
    work = WorkMeta("Лабораторная работа №3. Процессы", 3, "os")
    assert resolve_to("student.full_name", student, work) == "А А А"
    assert resolve_to("work.title", student, work) == "Лабораторная работа №3. Процессы"


def test_load_title_fields_confirmed(tmp_path: Path):
    p = tmp_path / "title-fields.yaml"
    p.write_text(
        yaml.safe_dump(
            {
                "fields": {
                    "student_name": {"find": "Петров", "to": "student.full_name"},
                    "topic": {"find": "Тема", "to": "work.title"},
                },
                "body_start": {"type": "heading_text", "value": "Цель работы"},
                "figure_caption": "Рисунок {n} – {caption}",
                "table_caption": "Таблица {n} – {caption}",
                "confirmed": True,
            },
            allow_unicode=True,
        ),
        encoding="utf-8",
    )
    tf = load_title_fields(p)
    assert tf.confirmed is True
    assert tf.body_start.type == "heading_text"
    assert tf.fields["student_name"].find == "Петров"


def test_load_title_fields_missing_body_start(tmp_path: Path):
    p = tmp_path / "title-fields.yaml"
    p.write_text("fields: {}\nconfirmed: true\n", encoding="utf-8")
    with pytest.raises(ValueError, match="body_start"):
        load_title_fields(p)
