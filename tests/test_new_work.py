from pathlib import Path
import hashlib
import pytest
from tools.new_work import NewWorkError, create_work


def _sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def test_create_work_copies_sample_and_does_not_mutate_it(tmp_path: Path):
    sample = tmp_path / "subjects" / "os" / "sample.docx"
    sample.parent.mkdir(parents=True)
    sample.write_bytes(b"PK-SAMPLE")
    before = _sha(sample)
    pdf = tmp_path / "task.pdf"
    pdf.write_bytes(b"%PDF")
    work = create_work("os", "processes", pdf, number=1, repo=tmp_path)
    assert work == tmp_path / "subjects" / "os" / "lab-01-processes"
    assert (work / "report.docx").read_bytes() == b"PK-SAMPLE"
    assert (work / "assignment.pdf").read_bytes() == b"%PDF"
    assert _sha(sample) == before


def test_create_work_missing_sample(tmp_path: Path):
    (tmp_path / "subjects" / "os").mkdir(parents=True)
    with pytest.raises(NewWorkError, match="Нет эталона"):
        create_work("os", "x", number=1, repo=tmp_path)


def test_second_call_does_not_clobber_report_md(tmp_path: Path):
    sample = tmp_path / "subjects" / "os" / "sample.docx"
    sample.parent.mkdir(parents=True)
    sample.write_bytes(b"PK-SAMPLE")
    work = create_work("os", "processes", number=1, repo=tmp_path)
    (work / "report.md").write_text("keep", encoding="utf-8")
    (work / "report.docx").write_bytes(b"DIRTY")
    create_work("os", "processes", number=1, repo=tmp_path)
    assert (work / "report.md").read_text(encoding="utf-8") == "keep"
    assert (work / "report.docx").read_bytes() == b"DIRTY"


def test_same_practice_second_author_gets_own_folder(tmp_path: Path):
    sample = tmp_path / "subjects" / "os" / "sample.docx"
    sample.parent.mkdir(parents=True)
    sample.write_bytes(b"PK-SAMPLE")
    first = tmp_path / "student-a.yaml"
    first.write_text("full_name: A\nname_short: A\ngroup: G\n", encoding="utf-8")
    second = tmp_path / "student-b.yaml"
    second.write_text("full_name: B\nname_short: B\ngroup: G\n", encoding="utf-8")
    w1 = create_work("os", "pz1-ivanov", number=1, student=first, repo=tmp_path)
    w2 = create_work("os", "pz1-petrov", number=1, student=second, repo=tmp_path)
    assert w1 != w2
    assert w1.name.startswith("lab-01-")
    assert w2.name.startswith("lab-01-")
    assert (w1 / "student.yaml").read_text(encoding="utf-8").startswith("full_name: A")
    assert (w2 / "student.yaml").read_text(encoding="utf-8").startswith("full_name: B")
    assert (w1 / "report.docx").is_file()
    assert (w2 / "report.docx").is_file()


def test_reset_docx_recopies_sample(tmp_path: Path):
    sample = tmp_path / "subjects" / "os" / "sample.docx"
    sample.parent.mkdir(parents=True)
    sample.write_bytes(b"PK-SAMPLE")
    work = create_work("os", "processes", number=1, repo=tmp_path)
    (work / "report.docx").write_bytes(b"DIRTY")
    create_work("os", "processes", number=1, repo=tmp_path, reset_docx=True)
    assert (work / "report.docx").read_bytes() == b"PK-SAMPLE"
