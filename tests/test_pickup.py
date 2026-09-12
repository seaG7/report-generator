from pathlib import Path
from dataclasses import dataclass
import hashlib
import pytest
from tools.pickup import (
    PickupError,
    claim_bundle,
    claim_pending,
    list_pending,
    preview_pending,
)


def test_list_pending_finds_inbox_and_loose_pdf(tmp_path: Path):
    subj = tmp_path / "subjects" / "os"
    inbox = subj / "inbox"
    inbox.mkdir(parents=True)
    (inbox / "lr3-processes.pdf").write_bytes(b"%PDF-inbox")
    (subj / "extra.pdf").write_bytes(b"%PDF-loose")
    lab = subj / "lab-01-old"
    lab.mkdir()
    (lab / "assignment.pdf").write_bytes(b"%PDF-already")
    pending = list_pending(tmp_path)
    paths = sorted(p.pdf.name for p in pending)
    assert paths == ["extra.pdf", "lr3-processes.pdf"]
    assert {p.subject for p in pending} == {"os"}


def test_claim_pending_moves_pdf_into_new_lab(tmp_path: Path):
    subj = tmp_path / "subjects" / "os"
    inbox = subj / "inbox"
    inbox.mkdir(parents=True)
    sample = subj / "sample.docx"
    sample.write_bytes(b"PK-SAMPLE")
    before = hashlib.sha256(sample.read_bytes()).hexdigest()
    pdf = inbox / "LR-3 processes.pdf"
    pdf.write_bytes(b"%PDF-inbox")
    pending = list_pending(tmp_path)
    assert len(pending) == 1
    work = claim_pending(pending[0], repo=tmp_path)
    assert work.name.startswith("lab-01-")
    assert (work / "assignment.pdf").read_bytes() == b"%PDF-inbox"
    assert (work / "report.docx").read_bytes() == b"PK-SAMPLE"
    assert not pdf.exists()
    assert list_pending(tmp_path) == []
    assert hashlib.sha256(sample.read_bytes()).hexdigest() == before


def test_claim_without_sample_raises(tmp_path: Path):
    subj = tmp_path / "subjects" / "os"
    inbox = subj / "inbox"
    inbox.mkdir(parents=True)
    (inbox / "a.pdf").write_bytes(b"%PDF")
    pending = list_pending(tmp_path)
    with pytest.raises(PickupError, match="Нет эталона"):
        claim_pending(pending[0], repo=tmp_path)


def test_claim_two_pdfs_increments_lab_number(tmp_path: Path):
    subj = tmp_path / "subjects" / "os"
    inbox = subj / "inbox"
    inbox.mkdir(parents=True)
    (subj / "sample.docx").write_bytes(b"PK")
    (inbox / "one.pdf").write_bytes(b"%PDF1")
    (inbox / "two.pdf").write_bytes(b"%PDF2")
    pending = list_pending(tmp_path)
    w1 = claim_pending(pending[0], repo=tmp_path)
    w2 = claim_pending(list_pending(tmp_path)[0], repo=tmp_path)
    names = sorted([w1.name, w2.name])
    assert names[0].startswith("lab-01-")
    assert names[1].startswith("lab-02-")


@dataclass
class _FakeConvert:
    md_path: Path
    image_paths: list
    warning: None
    empty_text: bool


def _fake_convert(pdf: Path, out_dir: Path) -> _FakeConvert:
    out_dir.mkdir(parents=True, exist_ok=True)
    md = out_dir / "document.md"
    md.write_text(f"PREVIEW {pdf.name}", encoding="utf-8")
    return _FakeConvert(md, [], None, False)


def test_preview_pending_does_not_move_pdfs(tmp_path: Path):
    subj = tmp_path / "subjects" / "os"
    inbox = subj / "inbox"
    inbox.mkdir(parents=True)
    pdf = inbox / "task.pdf"
    pdf.write_bytes(b"%PDF")
    previews = preview_pending(tmp_path, convert=_fake_convert)
    assert len(previews) == 1
    assert previews[0].pdf == pdf
    assert previews[0].md_path.is_file()
    assert "task.pdf" in previews[0].md_path.read_text(encoding="utf-8")
    assert pdf.is_file()
    assert list_pending(tmp_path)[0].pdf == pdf


def test_claim_bundle_keeps_lectures_out_of_new_labs(tmp_path: Path):
    subj = tmp_path / "subjects" / "os"
    inbox = subj / "inbox"
    inbox.mkdir(parents=True)
    (subj / "sample.docx").write_bytes(b"PK")
    task = inbox / "lab-3.pdf"
    lec = inbox / "lecture-05.pdf"
    extra = inbox / "syllabus.pdf"
    task.write_bytes(b"%PDF-task")
    lec.write_bytes(b"%PDF-lec")
    extra.write_bytes(b"%PDF-syl")
    work = claim_bundle(
        "os",
        assignment_pdfs=[task],
        lecture_pdfs=[lec],
        other_pdfs=[extra],
        repo=tmp_path,
    )
    assert work.name.startswith("lab-01-")
    assert (work / "assignment.pdf").read_bytes() == b"%PDF-task"
    assert (subj / "lectures" / "lecture-05.pdf").read_bytes() == b"%PDF-lec"
    assert (work / "lectures" / "lecture-05.pdf").is_file()
    assert (subj / "materials" / "syllabus.pdf").read_bytes() == b"%PDF-syl"
    assert not task.exists() and not lec.exists() and not extra.exists()
    assert list_pending(tmp_path) == []


def test_claim_bundle_two_assignment_parts_one_lab(tmp_path: Path):
    subj = tmp_path / "subjects" / "os"
    inbox = subj / "inbox"
    inbox.mkdir(parents=True)
    (subj / "sample.docx").write_bytes(b"PK")
    a1 = inbox / "part1.pdf"
    a2 = inbox / "part2.pdf"
    a1.write_bytes(b"%PDF1")
    a2.write_bytes(b"%PDF2")
    work = claim_bundle("os", assignment_pdfs=[a1, a2], repo=tmp_path)
    assert (work / "assignment.pdf").read_bytes() == b"%PDF1"
    extras = list(work.glob("assignment-*.pdf"))
    assert len(extras) == 1
    assert extras[0].read_bytes() == b"%PDF2"
    assert list_pending(tmp_path) == []
