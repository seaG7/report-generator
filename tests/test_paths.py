from pathlib import Path
import pytest
from tools.paths import SampleWriteError, assert_not_sample, is_sample_path, repo_root


def test_repo_root_finds_example(tmp_path, monkeypatch):
    assert (repo_root() / "student.yaml.example").is_file()


def test_is_sample_path_true_for_subject_sample(tmp_path):
    sample = tmp_path / "subjects" / "os" / "sample.docx"
    sample.parent.mkdir(parents=True)
    sample.write_bytes(b"PK")
    assert is_sample_path(sample) is True


def test_is_sample_path_false_for_work_copy(tmp_path):
    copy = tmp_path / "subjects" / "os" / "lab-01-x" / "report.docx"
    copy.parent.mkdir(parents=True)
    copy.write_bytes(b"PK")
    assert is_sample_path(copy) is False


def test_assert_not_sample_raises(tmp_path):
    sample = tmp_path / "subjects" / "os" / "sample.docx"
    sample.parent.mkdir(parents=True)
    sample.write_bytes(b"PK")
    with pytest.raises(SampleWriteError, match="Нельзя записывать эталон"):
        assert_not_sample(sample)
