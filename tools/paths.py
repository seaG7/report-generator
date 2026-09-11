from pathlib import Path


class SampleWriteError(Exception):
    pass


def repo_root() -> Path:
    cur = Path(__file__).resolve().parent
    for candidate in (cur, *cur.parents):
        if (candidate / "student.yaml.example").is_file():
            return candidate
    raise FileNotFoundError("Не найден корень репозитория (нет student.yaml.example)")


def is_sample_path(path: Path) -> bool:
    resolved = path.resolve()
    parts = resolved.parts
    if resolved.name != "sample.docx" or "subjects" not in parts:
        return False
    idx = parts.index("subjects")
    return idx == len(parts) - 3  # ..., subjects, <slug>, sample.docx


def assert_not_sample(path: Path) -> None:
    if is_sample_path(path):
        raise SampleWriteError("Нельзя записывать эталон sample.docx")
