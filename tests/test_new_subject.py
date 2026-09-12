from pathlib import Path
from tools.new_subject import new_subject


def test_new_subject_creates_folder_without_inbox(tmp_path: Path):
    d = new_subject("os", repo=tmp_path)
    assert d == tmp_path / "subjects" / "os"
    assert d.is_dir()
    assert not (d / "inbox").exists()
    assert not (d / "sample.docx").exists()
