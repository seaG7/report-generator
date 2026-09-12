from pathlib import Path
import hashlib
import pytest
from tools.pickup import PickupError, claim_pending, list_pending


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
