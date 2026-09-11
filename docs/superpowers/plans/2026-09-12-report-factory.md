# Фабрика учебных отчётов Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Собрать в репозитории локальную фабрику: PDF задания → MD с картинками, копия эталонного `.docx`, правка титульника/тела клонами блоков, скриншоты/диаграммы с манифестом, рендер страниц и проектные скиллы Grok `/report`.

**Architecture:** Скрипты в `tools/` — единственные писатели `.docx`. Эталон `subjects/<slug>/sample.docx` только читается; каждая работа живёт в `lab-NN-<slug>/` и стартует с копии. Grok оркестрирует фазы через скиллы, сам Word не генерирует и в Composio Word не ходит.

**Tech Stack:** Python 3.11+, pytest, pymupdf + pymupdf4llm, python-docx, PyYAML, Pillow, reportlab (только фикстуры тестов). Рендер: Word COM (`pywin32`) на Windows, иначе LibreOffice. Диаграммы: `@mermaid-js/mermaid-cli` через npx, опционально.

**Spec:** `docs/superpowers/specs/2026-09-11-report-factory-design.md`

## Global Constraints

- Ошибки CLI — **на русском** в stderr; код 0 только при полном успехе.
- Нельзя записывать `subjects/*/sample.docx` (проверка `assert_not_sample`).
- Не создавать отчёт через `Document()` с нуля; только копия эталона + клоны блоков / replace текста.
- Не добавлять `marker` в `requirements.txt` (опциональный fallback).
- Имена работ: `subjects/<subject>/lab-<NN>-<slug>/` с ведущим нулём (`lab-01-processes`).
- `student.yaml` в `.gitignore`; в репо только `student.yaml.example`.
- Тесты не требуют Grok. `render_docx` skip, если нет Word и LibreOffice. `diagram` тестируется с подменным runner.

---

## File map

| Path | Responsibility |
|---|---|
| `tools/paths.py` | `repo_root`, `is_sample_path`, `assert_not_sample`, `SampleWriteError` |
| `tools/config.py` | `Student`, `WorkMeta`, `TitleField`, `BodyStart`, `TitleFields`, загрузка YAML, `resolve_to` |
| `tools/pdf_to_md.py` | PDF → `assignment.md` + `images/` |
| `tools/new_work.py` | папка работы, копия эталона → `report.docx` |
| `tools/patch_title.py` | атомарные find→to на копии |
| `tools/screenshot.py` | PNG + `manifest.json` (`ok: false`) |
| `tools/diagram.py` | mermaid → PNG |
| `tools/parse_report.py` | `report.md` → блоки |
| `tools/insert_figure.py` | клон drawing+подписи, media, extent |
| `tools/replace_body.py` | очистка тела, вставка клонов по блокам MD |
| `tools/render_docx.py` | docx → `preview/page-NN.png` |
| `tools/__init__.py` | пакет |
| `.grok/skills/report/SKILL.md` | оркестратор |
| `.grok/skills/pdf-to-md/SKILL.md` | ingest |
| `.grok/skills/build-report/SKILL.md` | compose+verify |
| `AGENTS.md` | инварианты проекта |
| `tests/fixtures_util.py` | сборка тестового PDF и sample.docx |
| `tests/test_*.py` | pytest |

---

### Task 1: Каркас репозитория

**Files:**
- Create: `requirements.txt`
- Create: `pytest.ini`
- Create: `.gitignore`
- Create: `student.yaml.example`
- Create: `AGENTS.md`
- Create: `subjects/.gitkeep`
- Create: `tools/__init__.py`
- Create: `tests/test_scaffold.py`

**Interfaces:**
- Consumes: ничего
- Produces: каркас, на который опираются все следующие задачи; ключи `student.yaml.example`: `full_name`, `name_short`, `group`, `faculty`, `university`, `city`, `year`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_scaffold.py
from pathlib import Path
import yaml

ROOT = Path(__file__).resolve().parents[1]


def test_student_example_has_required_keys():
    data = yaml.safe_load((ROOT / "student.yaml.example").read_text(encoding="utf-8"))
    for key in ("full_name", "name_short", "group", "faculty", "university", "city", "year"):
        assert key in data
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_scaffold.py::test_student_example_has_required_keys -v`

Expected: FAIL (file not found)

- [ ] **Step 3: Write files**

`requirements.txt`:

```
pymupdf>=1.24.0
pymupdf4llm>=0.0.17
python-docx>=1.1.0
PyYAML>=6.0
Pillow>=10.0
pytest>=8.0
reportlab>=4.0
pywin32>=306; sys_platform == "win32"
```

`pytest.ini`:

```ini
[pytest]
pythonpath = .
testpaths = tests
```

`.gitignore`:

```
student.yaml
subjects/**/work/
subjects/**/screenshots/
subjects/**/preview/
__pycache__/
.pytest_cache/
*.pyc
.venv/
venv/
```

`student.yaml.example`:

```yaml
full_name: "Иванов Иван Иванович"
name_short: "Иванов И.И."
group: "ИКБО-21-23"
faculty: ""
university: ""
city: ""
year: 2026
```

`AGENTS.md`:

```markdown
# report-generator

- Эталон `subjects/<предмет>/sample.docx` только читать. Писать только в копию `lab-*/report.docx`.
- Не собирать Word с нуля и не восстанавливать стили списком правил. Копия эталона + клоны блоков.
- Не пропускать визуальный осмотр PNG (скрины, диаграммы, preview страниц).
- Не использовать Composio/HTML-to-Word для отчёта. Готовый `.docx` локальный.
- Диаграммы: Mermaid, не image generation, если хватает кода.
- CLI-ошибки скриптов на русском. Оркестрация: скилл `/report`.
```

Пустые `tools/__init__.py` и `subjects/.gitkeep`.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_scaffold.py -v`

Expected: PASS (нужен PyYAML: `pip install -r requirements.txt`)

- [ ] **Step 5: Commit**

```
git add requirements.txt pytest.ini .gitignore student.yaml.example AGENTS.md subjects/.gitkeep tools/__init__.py tests/test_scaffold.py
git commit -m "chore: scaffold report factory repo"
```

---

### Task 2: Защита эталона (`paths.py`)

**Files:**
- Create: `tools/paths.py`
- Test: `tests/test_paths.py`

**Interfaces:**
- Consumes: структура `subjects/<slug>/sample.docx`
- Produces:
  - `class SampleWriteError(Exception)`
  - `def repo_root() -> Path`
  - `def is_sample_path(path: Path) -> bool`
  - `def assert_not_sample(path: Path) -> None` — бросает `SampleWriteError` с текстом `Нельзя записывать эталон sample.docx`

`repo_root()`: от `Path(__file__).resolve()` идти вверх, пока не найдётся `student.yaml.example`.

`is_sample_path`: `path.resolve().name == "sample.docx"` и `"subjects"` есть в `parts`, и родительская папка — slug предмета (родитель родителя называется `subjects`).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_paths.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_paths.py -v`

Expected: FAIL `ModuleNotFoundError: tools.paths`

- [ ] **Step 3: Write minimal implementation**

```python
# tools/paths.py
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_paths.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/paths.py tests/test_paths.py
git commit -m "feat: block writes to subject sample.docx"
```

---

### Task 3: Конфиги YAML

**Files:**
- Create: `tools/config.py`
- Test: `tests/test_config.py`

**Interfaces:**
- Consumes: форматы из спеки
- Produces:

```python
from dataclasses import dataclass, field

@dataclass
class Student:
    full_name: str
    name_short: str
    group: str
    faculty: str = ""
    university: str = ""
    city: str = ""
    year: int = 2026

@dataclass
class WorkMeta:
    title: str
    number: int
    subject: str

@dataclass
class TitleField:
    find: str
    to: str  # "student.full_name" | "work.title" | ...

@dataclass
class BodyStart:
    type: str  # heading_text | paragraph_index | section
    value: str | int

@dataclass
class TitleFields:
    fields: dict[str, TitleField]
    body_start: BodyStart
    figure_caption: str
    table_caption: str
    confirmed: bool

def load_student(path: Path) -> Student: ...
def load_work(path: Path) -> WorkMeta: ...
def load_title_fields(path: Path) -> TitleFields: ...
def resolve_to(to: str, student: Student, work: WorkMeta) -> str: ...
```

`resolve_to("student.full_name", ...)` → `student.full_name`. Неизвестный префикс → `ValueError` с русским текстом. Пустой атрибут студента возвращает `""`.

`load_title_fields`: нет ключа `confirmed` → `confirmed=False`. Нет `body_start` → `ValueError("В title-fields.yaml нет body_start")`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_config.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_config.py -v`

Expected: FAIL import `tools.config`

- [ ] **Step 3: Write `tools/config.py`**

```python
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class Student:
    full_name: str
    name_short: str
    group: str
    faculty: str = ""
    university: str = ""
    city: str = ""
    year: int = 2026


@dataclass
class WorkMeta:
    title: str
    number: int
    subject: str


@dataclass
class TitleField:
    find: str
    to: str


@dataclass
class BodyStart:
    type: str
    value: str | int


@dataclass
class TitleFields:
    fields: dict[str, TitleField]
    body_start: BodyStart
    figure_caption: str
    table_caption: str
    confirmed: bool


def load_student(path: Path) -> Student:
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    return Student(
        full_name=data["full_name"],
        name_short=data["name_short"],
        group=data["group"],
        faculty=data.get("faculty") or "",
        university=data.get("university") or "",
        city=data.get("city") or "",
        year=int(data.get("year") or 2026),
    )


def load_work(path: Path) -> WorkMeta:
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    return WorkMeta(title=data["title"], number=int(data["number"]), subject=data["subject"])


def load_title_fields(path: Path) -> TitleFields:
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    if "body_start" not in data:
        raise ValueError("В title-fields.yaml нет body_start")
    bs = data["body_start"]
    btype = bs["type"]
    bval: Any = bs["value"]
    if btype in ("paragraph_index", "section"):
        bval = int(bval)
    fields = {
        name: TitleField(find=item["find"], to=item["to"])
        for name, item in (data.get("fields") or {}).items()
    }
    return TitleFields(
        fields=fields,
        body_start=BodyStart(type=btype, value=bval),
        figure_caption=data.get("figure_caption") or "Рисунок {n} – {caption}",
        table_caption=data.get("table_caption") or "Таблица {n} – {caption}",
        confirmed=bool(data.get("confirmed", False)),
    )


def resolve_to(to: str, student: Student, work: WorkMeta) -> str:
    if "." not in to:
        raise ValueError(f"Некорректный to: {to}")
    prefix, attr = to.split(".", 1)
    if prefix == "student":
        return str(getattr(student, attr))
    if prefix == "work":
        return str(getattr(work, attr))
    raise ValueError(f"Неизвестный префикс to: {to}")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_config.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/config.py tests/test_config.py
git commit -m "feat: load student, work, and title-fields yaml"
```

---

### Task 4: PDF → Markdown с картинками в порядке

**Files:**
- Create: `tools/pdf_to_md.py`
- Create: `tests/fixtures_util.py`
- Test: `tests/test_pdf_to_md.py`

**Interfaces:**
- Consumes: `pymupdf4llm.to_markdown`
- Produces:

```python
@dataclass
class ConvertResult:
    md_path: Path
    image_paths: list[Path]
    warning: str | None
    empty_text: bool

def convert_pdf(pdf_path: Path, out_dir: Path) -> ConvertResult:
    """Пишет out_dir/assignment.md и out_dir/images/*.
    empty_text=True если после удаления markdown-картинок и заголовков осталось < 40 символов.
    Если empty_text — пробует shutil.which("marker") или python -m marker; если нет — ConvertError.
    warning про число картинок PDF vs файлов в images/.
    """

class ConvertError(Exception):
    pass
```

CLI: `python tools/pdf_to_md.py PDF --out DIR` → код 1 и «Нужен текстовый PDF или Marker», если пустой текст и нет Marker.

`tests/fixtures_util.py` — функция `make_two_image_pdf(path: Path) -> None`: reportlab A4, текст `BEFORE`, красный PNG 40×40, текст `MIDDLE`, синий PNG 40×40, текст `AFTER`. Красный и синий сохранить рядом как `red.png`/`blue.png` во временной папке.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_pdf_to_md.py
from pathlib import Path
from tests.fixtures_util import make_two_image_pdf
from tools.pdf_to_md import convert_pdf


def test_images_appear_in_document_order(tmp_path: Path):
    pdf = tmp_path / "a.pdf"
    make_two_image_pdf(pdf)
    out = tmp_path / "out"
    result = convert_pdf(pdf, out)
    md = result.md_path.read_text(encoding="utf-8")
    assert result.md_path == out / "assignment.md"
    assert (out / "images").is_dir()
    assert len(result.image_paths) >= 2
    # порядок ссылок в MD = порядок BEFORE/red, затем blue, относительно текста
    pos_before = md.find("BEFORE")
    pos_middle = md.find("MIDDLE")
    pos_after = md.find("AFTER")
    import re
    imgs = [m.start() for m in re.finditer(r"!\[[^\]]*\]\([^)]+\)", md)]
    assert pos_before != -1 and pos_middle != -1 and pos_after != -1
    assert len(imgs) >= 2
    assert pos_before < imgs[0] < pos_middle < imgs[1] < pos_after
    for p in result.image_paths:
        assert p.is_file() and p.stat().st_size > 0
```

`make_two_image_pdf` в `tests/fixtures_util.py`:

```python
from pathlib import Path
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def make_two_image_pdf(path: Path) -> None:
    path = Path(path)
    img_dir = path.parent / "_pdf_imgs"
    img_dir.mkdir(parents=True, exist_ok=True)
    red = img_dir / "red.png"
    blue = img_dir / "blue.png"
    Image.new("RGB", (40, 40), (255, 0, 0)).save(red)
    Image.new("RGB", (40, 40), (0, 0, 255)).save(blue)
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFont("Times-Roman", 14)
    c.drawString(72, h - 72, "BEFORE")
    c.drawImage(str(red), 72, h - 160, 40, 40)
    c.drawString(72, h - 200, "MIDDLE")
    c.drawImage(str(blue), 72, h - 280, 40, 40)
    c.drawString(72, h - 320, "AFTER")
    c.save()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_pdf_to_md.py::test_images_appear_in_document_order -v`

Expected: FAIL import `tools.pdf_to_md`

- [ ] **Step 3: Write `tools/pdf_to_md.py`**

Алгоритм `convert_pdf`:

1. `out_dir.mkdir(parents=True)`, `images = out_dir / "images"`, `images.mkdir`.
2. `import pymupdf4llm, pymupdf`.
3. `md = pymupdf4llm.to_markdown(str(pdf_path), write_images=True, image_path=str(images))`.
4. Переписать абсолютные пути картинок в MD на относительные `images/<name>`.
5. Собрать `image_paths` = отсортированные png/jpg в `images/` в порядке первого появления в MD (не по имени файла).
6. Посчитать `len(doc.get_page_images())` по всем страницам pymupdf; если больше, чем файлов — `warning = "В PDF больше изображений, чем извлечено: {pdf_n} vs {disk_n}"`.
7. Текст без картинок: regex удалить `!\[.*?\]\(.*?\)`, strip; `empty_text = len(stripped) < 40`.
8. Если `empty_text`: если `shutil.which("marker")` — вызвать `marker_single` в `out_dir`; иначе `raise ConvertError("Нужен текстовый PDF или Marker")`.
9. Записать `out_dir / "assignment.md"`.
10. CLI: argparse, `ConvertError` → stderr + exit 1.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_pdf_to_md.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/pdf_to_md.py tests/fixtures_util.py tests/test_pdf_to_md.py
git commit -m "feat: convert assignment PDF to markdown with ordered images"
```

---

### Task 5: Новая работа и копия эталона

**Files:**
- Create: `tools/new_work.py`
- Test: `tests/test_new_work.py`

**Interfaces:**
- Consumes: `assert_not_sample`, `repo_root`
- Produces:

```python
class NewWorkError(Exception):
    pass

def create_work(
    subject: str,
    slug: str,
    pdf: Path | None = None,
    *,
    number: int | None = None,
    reset_docx: bool = False,
    repo: Path | None = None,
) -> Path:
    """
    repo/subjects/<subject>/lab-<NN>-<slug>/
    Копирует sample.docx → report.docx, если report.docx нет или reset_docx=True.
    Не трогает work/, screenshots/, report.md.
    Копирует pdf → assignment.pdf, если pdf задан и assignment.pdf ещё нет.
    Нет sample.docx → NewWorkError("Нет эталона subjects/<subject>/sample.docx").
    """
```

`NN`: если `number` задан — он; иначе max существующих `lab-(\d+)-` в предмете + 1, стартуя с 1. Формат `{n:02d}`.

Копия: `shutil.copy2(sample, dest)` после `assert_not_sample(dest)` (dest — report.docx, не sample). Эталон открывать только на чтение (`open(sample, "rb")` + copyfileobj или copy2).

CLI: `--subject --slug [--pdf] [--number] [--reset-docx] [--repo]`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_new_work.py
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


def test_reset_docx_recopies_sample(tmp_path: Path):
    sample = tmp_path / "subjects" / "os" / "sample.docx"
    sample.parent.mkdir(parents=True)
    sample.write_bytes(b"PK-SAMPLE")
    work = create_work("os", "processes", number=1, repo=tmp_path)
    (work / "report.docx").write_bytes(b"DIRTY")
    create_work("os", "processes", number=1, repo=tmp_path, reset_docx=True)
    assert (work / "report.docx").read_bytes() == b"PK-SAMPLE"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_new_work.py -v`

Expected: FAIL import

- [ ] **Step 3: Implement `create_work` and CLI**

```python
# tools/new_work.py
from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path

from tools.paths import assert_not_sample, repo_root


class NewWorkError(Exception):
    pass


_LAB = re.compile(r"^lab-(\d+)-")


def _next_number(subject_dir: Path) -> int:
    nums = []
    if subject_dir.is_dir():
        for p in subject_dir.iterdir():
            m = _LAB.match(p.name)
            if m:
                nums.append(int(m.group(1)))
    return (max(nums) + 1) if nums else 1


def create_work(
    subject: str,
    slug: str,
    pdf: Path | None = None,
    *,
    number: int | None = None,
    reset_docx: bool = False,
    repo: Path | None = None,
) -> Path:
    root = Path(repo) if repo is not None else repo_root()
    subject_dir = root / "subjects" / subject
    sample = subject_dir / "sample.docx"
    if not sample.is_file():
        raise NewWorkError(f"Нет эталона subjects/{subject}/sample.docx")
    n = number if number is not None else _next_number(subject_dir)
    work = subject_dir / f"lab-{n:02d}-{slug}"
    work.mkdir(parents=True, exist_ok=True)
    dest = work / "report.docx"
    assert_not_sample(dest)
    if reset_docx or not dest.is_file():
        shutil.copy2(sample, dest)
    if pdf is not None:
        target_pdf = work / "assignment.pdf"
        if not target_pdf.is_file():
            shutil.copy2(pdf, target_pdf)
    return work


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--subject", required=True)
    p.add_argument("--slug", required=True)
    p.add_argument("--pdf")
    p.add_argument("--number", type=int)
    p.add_argument("--reset-docx", action="store_true")
    p.add_argument("--repo")
    args = p.parse_args()
    try:
        work = create_work(
            args.subject,
            args.slug,
            Path(args.pdf) if args.pdf else None,
            number=args.number,
            reset_docx=args.reset_docx,
            repo=Path(args.repo) if args.repo else None,
        )
    except NewWorkError as e:
        print(str(e), file=sys.stderr)
        return 1
    print(work)
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_new_work.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/new_work.py tests/test_new_work.py
git commit -m "feat: create lab folder and copy sample report"
```

---

### Task 6: Атомарная замена титульника

**Files:**
- Create: `tools/patch_title.py`
- Modify: `tests/fixtures_util.py` — добавить `make_sample_docx(path: Path, image: Path | None = None) -> None`
- Test: `tests/test_patch_title.py`

**Interfaces:**
- Consumes: `assert_not_sample`, `TitleFields`, `Student`, `WorkMeta`, `resolve_to`
- Produces:

```python
class TitleFieldNotFound(Exception):
    pass

class TitleNotConfirmed(Exception):
    pass

def patch_title(
    docx_path: Path,
    title_fields: TitleFields,
    student: Student,
    work: WorkMeta,
) -> None:
    """
    Если не title_fields.confirmed → TitleNotConfirmed("title-fields.yaml не подтверждён").
    Копия в *.tmp.docx, замены, затем replace. Любая find без вхождения →
    TitleFieldNotFound, исходный файл байт-в-байт тот же.
    assert_not_sample(docx_path) в начале.
    """
```

Замена в paragraph.text: если `find in paragraph.text`, склеить в первый run, остальные runs обнулить. То же для ячеек таблиц, header/footer.

`make_sample_docx`: python-docx, абзацы `Университет`, `Петров Пётр Петрович`, `ИКБО-00-00`, `Лабораторная работа №1. Пример`, page break, heading `Цель работы`, абзац `Текст цели.`, картинка 40×40 + абзац `Рисунок 1 – Пример`, абзац `Таблица 1 – Пример`, таблица 2×2 Table Grid (`A` `B` / `1` `2`), heading `Выводы`, абзац `Текст выводов.`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_patch_title.py
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
    texts = [p.text for p in Document(docx).paragraphs]
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
```

Дописать `make_sample_docx` в `tests/fixtures_util.py` (PIL 40×40, `add_picture`, `add_heading` level=1, `add_page_break`).

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_patch_title.py -v`

Expected: FAIL import `tools.patch_title`

- [ ] **Step 3: Implement atomic patch**

```python
def patch_title(docx_path, title_fields, student, work):
    assert_not_sample(docx_path)
    if not title_fields.confirmed:
        raise TitleNotConfirmed("title-fields.yaml не подтверждён")
    mapping = {}
    for field in title_fields.fields.values():
        mapping[field.find] = resolve_to(field.to, student, work)
    tmp = docx_path.with_suffix(".tmp.docx")
    shutil.copy2(docx_path, tmp)
    try:
        _apply_mapping(tmp, mapping)
        os.replace(tmp, docx_path)
    except Exception:
        if tmp.exists():
            tmp.unlink()
        raise
```

`_apply_mapping`: сначала проверить, что каждый `find` встречается хотя бы раз в документе (параграфы, таблицы, headers, footers). Если нет — `TitleFieldNotFound(f"Строка не найдена на титульнике: {find}")`. Затем заменить.

CLI: `--docx --title-fields --student --work`

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_patch_title.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/patch_title.py tests/fixtures_util.py tests/test_patch_title.py
git commit -m "feat: atomically patch title page fields on sample copy"
```

---

### Task 7: Скриншоты и манифест

**Files:**
- Create: `tools/screenshot.py`
- Test: `tests/test_screenshot.py`

**Interfaces:**
- Consumes: ничего из docx
- Produces:

```python
ALLOWED_SOURCES = ("browser", "terminal", "ide", "diagram", "other")

def register_image(
    src: Path,
    out_dir: Path,
    caption: str,
    source: str,
    slug: str | None = None,
) -> Path:
    """Копирует src → out_dir/{NN:02d}-{slug}.png, append manifest ok=False. Возвращает путь PNG."""

def capture_screen(out_dir: Path, caption: str, source: str, slug: str | None = None) -> Path:
    """Pillow ImageGrab.grab() → PNG, затем как register_image."""

def mark_ok(out_dir: Path, file_name: str) -> None:
    """Ставит ok=True у записи манифеста. Нет записи → ManifestError."""

def load_manifest(out_dir: Path) -> list[dict]: ...
```

`NN` = max существующих префиксов `^\d+` в dir + 1, старт 1. `source` не из ALLOWED_SOURCES → `ValueError("Неизвестный source")`.

Манифест: `out_dir / "manifest.json"`, список объектов `file`, `caption`, `source`, `ok`.

CLI: `python tools/screenshot.py --from-file IMG --out-dir DIR --caption "..." --source terminal [--slug install]`  
`python tools/screenshot.py --capture --out-dir DIR --caption "..." --source ide`  
`python tools/screenshot.py --mark-ok FILE --out-dir DIR`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_screenshot.py
from pathlib import Path
from PIL import Image
import json
import pytest
from tools.screenshot import ManifestError, mark_ok, register_image


def test_register_image_writes_manifest_ok_false(tmp_path: Path):
    src = tmp_path / "a.png"
    Image.new("RGB", (10, 10), (1, 2, 3)).save(src)
    out = tmp_path / "screenshots"
    dest = register_image(src, out, "Установка зависимостей", "terminal", slug="install")
    assert dest.name == "01-install.png"
    assert dest.is_file()
    data = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
    assert data == [{
        "file": "01-install.png",
        "caption": "Установка зависимостей",
        "source": "terminal",
        "ok": False,
    }]
    mark_ok(out, "01-install.png")
    data = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
    assert data[0]["ok"] is True


def test_second_image_increments_index(tmp_path: Path):
    src = tmp_path / "a.png"
    Image.new("RGB", (10, 10)).save(src)
    out = tmp_path / "screenshots"
    register_image(src, out, "a", "browser", slug="a")
    dest = register_image(src, out, "b", "browser", slug="b")
    assert dest.name == "02-b.png"


def test_mark_ok_missing_raises(tmp_path: Path):
    out = tmp_path / "screenshots"
    out.mkdir()
    (out / "manifest.json").write_text("[]", encoding="utf-8")
    with pytest.raises(ManifestError):
        mark_ok(out, "nope.png")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_screenshot.py -v`

Expected: FAIL import

- [ ] **Step 3: Implement screenshot.py**

```python
# tools/screenshot.py
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

ALLOWED_SOURCES = ("browser", "terminal", "ide", "diagram", "other")
_IDX = re.compile(r"^(\d+)-")


class ManifestError(Exception):
    pass


def load_manifest(out_dir: Path) -> list[dict]:
    path = Path(out_dir) / "manifest.json"
    if not path.is_file():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def _save_manifest(out_dir: Path, data: list[dict]) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "manifest.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _next_index(out_dir: Path) -> int:
    nums = []
    if out_dir.is_dir():
        for p in out_dir.iterdir():
            m = _IDX.match(p.name)
            if m:
                nums.append(int(m.group(1)))
    return (max(nums) + 1) if nums else 1


def register_image(
    src: Path,
    out_dir: Path,
    caption: str,
    source: str,
    slug: str | None = None,
) -> Path:
    if source not in ALLOWED_SOURCES:
        raise ValueError("Неизвестный source")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    n = _next_index(out_dir)
    slug = slug or "shot"
    dest = out_dir / f"{n:02d}-{slug}.png"
    shutil.copy2(src, dest)
    data = load_manifest(out_dir)
    data.append({"file": dest.name, "caption": caption, "source": source, "ok": False})
    _save_manifest(out_dir, data)
    return dest


def capture_screen(out_dir: Path, caption: str, source: str, slug: str | None = None) -> Path:
    from PIL import ImageGrab

    tmp = Path(out_dir) / "_capture.png"
    tmp.parent.mkdir(parents=True, exist_ok=True)
    ImageGrab.grab().save(tmp)
    try:
        return register_image(tmp, out_dir, caption, source, slug=slug)
    finally:
        tmp.unlink(missing_ok=True)


def mark_ok(out_dir: Path, file_name: str) -> None:
    data = load_manifest(out_dir)
    for item in data:
        if item["file"] == file_name:
            item["ok"] = True
            _save_manifest(Path(out_dir), data)
            return
    raise ManifestError(f"Нет записи в манифесте: {file_name}")
```

CLI: argparse с взаимоисключающими `--from-file` / `--capture` / `--mark-ok`.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_screenshot.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/screenshot.py tests/test_screenshot.py
git commit -m "feat: screenshot registry with ok flag manifest"
```

---

### Task 8: Mermaid → PNG

**Files:**
- Create: `tools/diagram.py`
- Test: `tests/test_diagram.py`

**Interfaces:**
- Consumes: внешний mermaid-cli
- Produces:

```python
class DiagramError(Exception):
    pass

def render_mermaid(
    mmd_path: Path,
    png_path: Path | None = None,
    *,
    runner=None,
) -> Path:
    """
    png_path по умолчанию: тот же stem, суффикс .png.
    runner(cmd: list[str]) -> int  (0 = ok). По умолчанию subprocess.run.
    Команда: npx --yes @mermaid-js/mermaid-cli -i <mmd> -o <png> -b white
    runner != 0 → DiagramError("Не удалось собрать диаграмму Mermaid").
    """
```

- [ ] **Step 1: Write the failing test**

```python
# tests/test_diagram.py
from pathlib import Path
import pytest
from tools.diagram import DiagramError, render_mermaid


def test_render_mermaid_uses_runner_and_writes_png(tmp_path: Path):
    mmd = tmp_path / "a.mmd"
    mmd.write_text("flowchart LR\nA-->B\n", encoding="utf-8")
    png = tmp_path / "a.png"

    def runner(cmd):
        assert cmd[0] == "npx"
        assert str(mmd) in cmd and str(png) in cmd
        png.write_bytes(b"\x89PNG")
        return 0

    out = render_mermaid(mmd, png, runner=runner)
    assert out == png and png.is_file()


def test_render_mermaid_nonzero_raises(tmp_path: Path):
    mmd = tmp_path / "a.mmd"
    mmd.write_text("x", encoding="utf-8")
    with pytest.raises(DiagramError, match="Mermaid"):
        render_mermaid(mmd, runner=lambda cmd: 1)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_diagram.py -v`

Expected: FAIL import

- [ ] **Step 3: Implement `render_mermaid` + CLI `python tools/diagram.py INPUT.mmd [--out OUT.png]`**

```python
# tools/diagram.py
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


class DiagramError(Exception):
    pass


def _default_runner(cmd: list[str]) -> int:
    return subprocess.run(cmd, check=False).returncode


def render_mermaid(
    mmd_path: Path,
    png_path: Path | None = None,
    *,
    runner=None,
) -> Path:
    mmd_path = Path(mmd_path)
    png_path = Path(png_path) if png_path else mmd_path.with_suffix(".png")
    run = runner or _default_runner
    cmd = [
        "npx",
        "--yes",
        "@mermaid-js/mermaid-cli",
        "-i",
        str(mmd_path),
        "-o",
        str(png_path),
        "-b",
        "white",
    ]
    if run(cmd) != 0:
        raise DiagramError("Не удалось собрать диаграмму Mermaid")
    return png_path
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_diagram.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/diagram.py tests/test_diagram.py
git commit -m "feat: render mermaid diagrams to png"
```

---

### Task 9: Парсер `report.md`

**Files:**
- Create: `tools/parse_report.py`
- Test: `tests/test_parse_report.py`

**Interfaces:**
- Consumes: конвенции спеки (heading, `![alt](path)`, `Таблица:`, GFM table, fenced code)
- Produces:

```python
@dataclass
class Block:
    kind: str  # heading | paragraph | figure | table | code
    text: str = ""
    level: int = 0          # heading 1-3
    path: str = ""          # figure
    caption: str = ""       # figure/table
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    lang: str = ""          # code

def parse_report_md(text: str) -> list[Block]:
```

Правила: `#` / `##` / `###`; картинка — строка целиком `![caption](path)`; строка `Таблица: NAME` привязывается к ближайшей следующей GFM-таблице; code fence ` ```lang `; пустые строки разделяют абзацы. Остальное — `paragraph`.

- [ ] **Step 1: Write the failing test**

````python
# tests/test_parse_report.py
from tools.parse_report import parse_report_md

MD = (
    "# Цель работы\n\n"
    "Текст цели.\n\n"
    "![Установка зависимостей](screenshots/01-install.png)\n\n"
    "## Ход\n\n"
    "Таблица: Результаты измерений\n\n"
    "| Параметр | Значение |\n"
    "| --- | --- |\n"
    "| t | 12 |\n\n"
    "```python\n"
    'print("hello")\n'
    "```\n"
)


def test_parse_report_md_kinds():
    blocks = parse_report_md(MD)
    kinds = [b.kind for b in blocks]
    assert kinds == ["heading", "paragraph", "figure", "heading", "table", "code"]
    assert blocks[0].level == 1 and blocks[0].text == "Цель работы"
    assert blocks[2].caption == "Установка зависимостей"
    assert blocks[2].path == "screenshots/01-install.png"
    assert blocks[4].caption == "Результаты измерений"
    assert blocks[4].headers == ["Параметр", "Значение"]
    assert blocks[4].rows == [["t", "12"]]
    assert blocks[5].lang == "python" and "print" in blocks[5].text
````

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_parse_report.py -v`

Expected: FAIL import

- [ ] **Step 3: Implement line-based parser (stdlib only)**

```python
# tools/parse_report.py
from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class Block:
    kind: str
    text: str = ""
    level: int = 0
    path: str = ""
    caption: str = ""
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    lang: str = ""


_IMG = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")
_TABLE_CAPTION = re.compile(r"^Таблица:\s*(.+)\s*$")
_HEADING = re.compile(r"^(#{1,3})\s+(.+?)\s*$")


def _split_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def parse_report_md(text: str) -> list[Block]:
    lines = text.replace("\r\n", "\n").split("\n")
    blocks: list[Block] = []
    i = 0
    pending_table_caption = ""
    para: list[str] = []

    def flush_para() -> None:
        nonlocal para
        body = "\n".join(para).strip()
        para = []
        if body:
            blocks.append(Block(kind="paragraph", text=body))

    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("```"):
            flush_para()
            lang = line.strip()[3:].strip()
            i += 1
            buf: list[str] = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            blocks.append(Block(kind="code", text="\n".join(buf), lang=lang))
            i += 1
            continue
        m_h = _HEADING.match(line)
        if m_h:
            flush_para()
            blocks.append(Block(kind="heading", text=m_h.group(2), level=len(m_h.group(1))))
            i += 1
            continue
        m_img = _IMG.match(line.strip())
        if m_img:
            flush_para()
            blocks.append(Block(kind="figure", caption=m_img.group(1), path=m_img.group(2)))
            i += 1
            continue
        m_tc = _TABLE_CAPTION.match(line.strip())
        if m_tc:
            flush_para()
            pending_table_caption = m_tc.group(1)
            i += 1
            continue
        if line.strip().startswith("|"):
            flush_para()
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(_split_row(lines[i]))
                i += 1
            if len(rows) >= 2 and set(rows[1][0]).issubset(set("-: ")):
                headers, data = rows[0], rows[2:]
            else:
                headers, data = (rows[0] if rows else []), rows[1:]
            blocks.append(
                Block(
                    kind="table",
                    caption=pending_table_caption,
                    headers=headers,
                    rows=data,
                )
            )
            pending_table_caption = ""
            continue
        if not line.strip():
            flush_para()
            i += 1
            continue
        para.append(line)
        i += 1
    flush_para()
    return blocks
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_parse_report.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/parse_report.py tests/test_parse_report.py
git commit -m "feat: parse report markdown into typed blocks"
```

---

### Task 10: Тело отчёта клонами блоков эталона

**Files:**
- Create: `tools/insert_figure.py`
- Create: `tools/replace_body.py`
- Test: `tests/test_replace_body.py`

**Interfaces:**
- Consumes: `parse_report_md`, `assert_not_sample`, `TitleFields`, `load_manifest`, `make_sample_docx`
- Produces:

```python
class ComposeError(Exception):
    pass

@dataclass
class Prototypes:
    figure_p: object | None
    figure_caption_p: object | None
    table: object | None
    table_caption_p: object | None
    body_style: str
    heading_styles: dict[int, str]
    code_p: object | None
    max_width_emu: int

def collect_prototypes(document) -> Prototypes:
    """Сканирует document до очистки тела. figure_p = первый w:p с w:drawing.
    Следующий параграф — caption. Первая таблица — table. Параграф сразу перед ней,
    если текст начинается с «Таблица» — table_caption_p.
    max_width_emu: ширина страницы минус left+right margin секции.
    """

def replace_body(
    docx_path: Path,
    report_md: Path,
    title_fields: TitleFields,
    work_dir: Path,
) -> None:
    """
    assert_not_sample. confirmed иначе ComposeError.
    Картинки из MD: файл должен существовать относительно work_dir.
    Если path начинается с screenshots/ (или лежит в work_dir/screenshots) —
    запись манифеста с ok=True обязательна, иначе ComposeError.
    Нет прототипа рисунка, а в MD есть figure → ComposeError("В эталоне нет образца рисунка").
    Аналогично таблица.
    body_start: heading_text — первый параграф с text.strip()==value (или startswith);
    paragraph_index — индекс; section — начало секции N.
    Удалить все параграфы/таблицы ПОСЛЕ body_start (body_start оставить и заменить
    его содержимое первым heading, если первый блок — тот же заголовок; иначе
    оставить body_start и вставлять после него). Проще и однозначно:
    удалить body_start и всё после него до sectPr, затем вставить все блоки MD.
    Титульник — всё ДО body_start — не трогать.
    """
```

`insert_figure.py`:

```python
EMU_PER_INCH = 914400

def image_extent_emus(png_path: Path, max_width_emu: int) -> tuple[int, int]:
    """Pillow size, cx = min(px_w/96*EMU, max_width_emu), cy пропорционально."""

def add_image_rel(document, png_path: Path) -> str:
    """document.part.get_or_add_image(str(png_path)); вернуть rId."""

def append_cloned_figure(document, proto_p, proto_cap, png_path: Path, caption: str, n: int, template: str, max_width_emu: int) -> None:
    """deepcopy proto_p и proto_cap в body перед sectPr. Обновить blip embed, wp:extent cx/cy,
    текст подписи = template.format(n=n, caption=caption). Все w:t в caption-параграфе
    заменить: первый run = полная строка, остальные пустые.
    """
```

Нумерация: отдельный счётчик figure и table, стартуя с 1, в порядке блоков MD.

Абзацы: `document.add_paragraph(text, style=prototypes.body_style)` допустимо — стиль из эталона, не новый Document.

Заголовки: style из `heading_styles[level]`, если ключа нет — body_style + bold run (только run, стили документа не создавать).

Таблица: deepcopy `w:tbl`, подогнать число строк/столбцов (если прототип меньше — клонировать последнюю tr/tc; если больше в прототипе — обрезать), заполнить текст ячеек. Перед таблицей caption-параграф по `table_caption` template.

Код: если `code_p` есть — clone и заменить текст; иначе абзац body_style, run `rFonts ascii=Consolas hAnsi=Consolas`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_replace_body.py
from pathlib import Path
from PIL import Image
from docx import Document
from tests.fixtures_util import make_sample_docx
from tools.config import BodyStart, Student, TitleField, TitleFields, WorkMeta
from tools.replace_body import ComposeError, replace_body
from tools.screenshot import mark_ok, register_image
import pytest

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
    doc = Document(docx)
    texts = [p.text for p in doc.paragraphs]
    joined = "\n".join(texts)
    assert "Петров Пётр Петрович" in joined  # титульник жив
    assert "Рисунок 1 – Первый" in joined
    assert "Рисунок 2 – Второй" in joined
    assert "Текст цели." not in joined  # старое тело ушло
    media = list((tmp_path / "lab").glob("**/*"))  # sanity
    # два изображения в word/media
    import zipfile
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
    # ok остаётся false
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
```

Обновить `make_sample_docx(path, image=None)`: если `image` None — сгенерировать серый 40×40 рядом с path.

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_replace_body.py -v`

Expected: FAIL import `tools.replace_body`

- [ ] **Step 3: Implement insert_figure.py and replace_body.py**

Ключевые XML-имена:

- drawing: `{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent` атрибуты `cx`, `cy`
- blip: `{http://schemas.openxmlformats.org/drawingml/2006/main}blip` атрибут `{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed`
- удаление тела: `body = document.element.body`; найти paragraph с `body_start`; все следующие sibling кроме `w:sectPr` — `remove`.

Вставка: `body.insert_element_before(elem, "w:sectPr")` если есть sectPr, иначе append.

CLI: `python tools/replace_body.py --docx --md --title-fields --work-dir`

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_replace_body.py -v`

Expected: PASS

- [ ] **Step 5: Commit**

```
git add tools/insert_figure.py tools/replace_body.py tests/test_replace_body.py tests/fixtures_util.py
git commit -m "feat: replace report body by cloning sample blocks"
```

---

### Task 11: Рендер страниц docx → PNG

**Files:**
- Create: `tools/render_docx.py`
- Test: `tests/test_render_docx.py`

**Interfaces:**
- Consumes: `assert_not_sample` не нужен (только читает docx), pymupdf для PDF→PNG
- Produces:

```python
class RenderUnavailable(Exception):
    pass

def render_docx(docx_path: Path, out_dir: Path) -> list[Path]:
    """
    out_dir/page-01.png ...
    1) Windows: win32com Word.Application Documents.Open, ExportAsFixedFormat PDF (17), Quit.
    2) Иначе soffice/libreoffice --headless --convert-to pdf --outdir <tmp>.
    3) PDF → PNG через pymupdf page.get_pixmap(dpi=150).
    Нет Word и нет бинаря soffice/libreoffice → RenderUnavailable("Нет Word и нет LibreOffice").
    Пустой PNG (stat().st_size == 0) не возвращать, бросать RenderUnavailable.
    """

def word_available() -> bool: ...
def soffice_available() -> bool: ...
```

Поиск soffice: `shutil.which("soffice")`, `shutil.which("libreoffice")`, типичные пути Windows `C:\Program Files\LibreOffice\program\soffice.exe`.

CLI: `python tools/render_docx.py DOCX --out-dir DIR`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_render_docx.py
from pathlib import Path
import pytest
from tests.fixtures_util import make_sample_docx
from tools.render_docx import RenderUnavailable, render_docx, soffice_available, word_available

pytestmark = pytest.mark.skipif(
    not word_available() and not soffice_available(),
    reason="Нет Word и нет LibreOffice",
)


def test_render_docx_writes_nonempty_png(tmp_path: Path):
    docx = tmp_path / "s.docx"
    make_sample_docx(docx)
    out = tmp_path / "preview"
    pages = render_docx(docx, out)
    assert pages, "ожидалась хотя бы одна страница"
    assert pages[0].name == "page-01.png"
    assert pages[0].stat().st_size > 0
```

Плюс тест без skip:

```python
def test_render_unavailable_message():
    from tools.render_docx import RenderUnavailable
    err = RenderUnavailable("Нет Word и нет LibreOffice")
    assert "LibreOffice" in str(err)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_render_docx.py -v`

Expected: FAIL import или skip + fail на `test_render_unavailable_message` если файл есть без класса — сначала FAIL import

- [ ] **Step 3: Implement Word COM then LibreOffice then pymupdf pixmaps**

Word ExportAsFixedFormat: второй аргумент `17` (`wdExportFormatPDF`). Всегда `word.Quit()` в `finally`. LibreOffice: `--convert-to pdf --outdir`. Имена `page-{n:02d}.png`.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_render_docx.py -v`

Expected: PASS (или skip render + PASS message test)

- [ ] **Step 5: Commit**

```
git add tools/render_docx.py tests/test_render_docx.py
git commit -m "feat: render docx pages to png for visual qa"
```

---

### Task 12: Скиллы Grok и README

**Files:**
- Create: `.grok/skills/report/SKILL.md`
- Create: `.grok/skills/pdf-to-md/SKILL.md`
- Create: `.grok/skills/build-report/SKILL.md`
- Modify: `AGENTS.md` — только если нужно добавить команды скриптов; не дублировать скилл
- Create: `README.md`
- Test: `tests/test_skills_exist.py`

**Interfaces:**
- Consumes: все CLI из задач 4–11
- Produces: скиллы, которые Grok вызывает по `/report`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_skills_exist.py
from pathlib import Path
from tools.paths import repo_root

def test_three_skills_have_names():
    root = repo_root() / ".grok" / "skills"
    for name in ("report", "pdf-to-md", "build-report"):
        text = (root / name / "SKILL.md").read_text(encoding="utf-8")
        assert text.startswith("---")
        assert f"name: {name}" in text
        assert "description:" in text
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_skills_exist.py -v`

Expected: FAIL file not found

- [ ] **Step 3: Write skills and README**

`.grok/skills/pdf-to-md/SKILL.md`:

```markdown
---
name: pdf-to-md
description: Convert an assignment PDF to Markdown with images in document order. Use when ingesting a lab PDF, converting PDF to md, or when /report needs assignment.md. Use when the user runs /pdf-to-md.
---

# PDF → Markdown

Run from repo root:

```
python tools/pdf_to_md.py "<path-to.pdf>" --out "subjects/<subject>/lab-NN-<slug>"
```

Read `assignment.md`. Do not invent figures missing from extraction. If the command exits 1 with «Нужен текстовый PDF или Marker», stop and tell the user.
```

`.grok/skills/build-report/SKILL.md`:

```markdown
---
name: build-report
description: Patch title fields and replace body of a copied sample.docx, then render pages for visual QA. Use when composing the Word report after work is done, or when the user runs /build-report.
---

# Сборка report.docx

Preconditions: `title-fields.yaml` has `confirmed: true`; `work.yaml` exists; `report.md` exists; every `screenshots/` image referenced in `report.md` has `ok: true` in `manifest.json`.

```
python tools/patch_title.py --docx "<work>/report.docx" --title-fields "subjects/<subject>/title-fields.yaml" --student student.yaml --work "<work>/work.yaml"
python tools/replace_body.py --docx "<work>/report.docx" --md "<work>/report.md" --title-fields "subjects/<subject>/title-fields.yaml" --work-dir "<work>"
python tools/render_docx.py "<work>/report.docx" --out-dir "<work>/preview"
```

Never write `sample.docx`. Never `Document()` from scratch. After render, open every `preview/page-*.png` and inspect title, figures, tables. On failure, fix the copy and re-render. After 3 failed inspect loops, stop and list defects. Do not declare the report ready without a passing inspect.
```

`.grok/skills/report/SKILL.md` — полный оркестратор:

Frontmatter:

```yaml
name: report
description: Run the full lab-report factory from assignment PDF plus subject sample.docx to a verified report.docx. Use when the user wants a lab report, /report, лабораторная, сделай отчёт, or attaches an assignment PDF with a formatted sample.
```

Тело скилла (вставить как есть):

1. Прочитать `student.yaml`. Нет — стоп, скопировать из `student.yaml.example`.
2. Определить предмет (аргумент, cwd, или один вопрос).
3. Нет `subjects/<subject>/sample.docx` — стоп, нужен эталон.
4. Нет `title-fields.yaml` или `confirmed` не true — онбординг: `python tools/new_work.py` не вызывать для лабы; скопировать эталон во временную папку только для рендера первой страницы `python tools/render_docx.py sample.docx --out-dir <tmp>/preview` (рендер **читает**, не пишет в sample). Прочитать page-01.png и текст титульника через unpack/`Document`. Предложить `title-fields.yaml` (find-строки, `body_start`, шаблоны подписей). Ждать подтверждения пользователя, затем записать файл с `confirmed: true`.
5. `python tools/new_work.py --subject S --slug SLUG [--pdf PDF] [--number N]`
6. Ingest: `python tools/pdf_to_md.py ...` прочитать assignment.md.
7. Execute в `work/`. Скрины: chrome-devtools или `python tools/screenshot.py --capture|--from-file`. Диаграммы: писать `.mmd`, `python tools/diagram.py`, затем `python tools/screenshot.py --from-file PNG --source diagram --out-dir screenshots --caption "..."`. **Открыть каждый PNG**, сверить с заданием, затем `python tools/screenshot.py --mark-ok FILE --out-dir screenshots`. Не ставить ok слепо. Не image_gen для блок-схем.
8. Написать `work.yaml` и `report.md` (структура как в эталоне; `![caption](screenshots/...)`; `Таблица: ...` перед GFM).
9. Вызвать шаги build-report (patch_title, replace_body, render_docx). Смотреть preview. Максимум 3 цикла правок. Не объявлять готовым без осмотра.
10. Не Composio Word. Эталон не модифицировать. «Пересобрать с эталона» = `python tools/new_work.py ... --reset-docx`.

`README.md`: что это, `pip install -r requirements.txt`, заполнить `student.yaml`, положить эталон в `subjects/<slug>/sample.docx`, в Grok `/report` + PDF. Таблица скриптов как в спеке. Требования Word или LibreOffice для verify.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_skills_exist.py tests/ -v --tb=short`

Expected: весь suite зелёный (кроме skip render без Word/LO)

- [ ] **Step 5: Commit**

```
git add .grok/skills/report/SKILL.md .grok/skills/pdf-to-md/SKILL.md .grok/skills/build-report/SKILL.md README.md tests/test_skills_exist.py AGENTS.md
git commit -m "feat: add Grok skills and README for /report factory"
```

---

## Notes for implementers

- Windows: команды `python -m pytest ...` и `git commit -m "..."` из корня `D:\Reps\report-generator`.
- `make_sample_docx` развивается в задачах 6 и 10 — не плодить вторую фабрику фикстур.
- Скилл `build-report` вызывает ровно три команды: `patch_title.py`, `replace_body.py`, `render_docx.py`.
- После Task 12 прогнать полный `python -m pytest`.

---

## Spec coverage (self-review)

| Спека | Задача |
|---|---|
| student.yaml / example / gitignore | 1 |
| sample.docx read-only `assert_not_sample` | 2, 5, 6, 10 |
| title-fields + confirmed + body_start | 3, 6, 10, 12 |
| work.yaml | 3, 6, 12 |
| PDF→MD порядок картинок, Marker fallback | 4 |
| lab-NN-slug, копия эталона, не затирать work/md | 5 |
| --reset-docx | 5, 12 |
| атомарный титульник, стоп если find нет | 6 |
| screenshots/manifest ok | 7, 10, 12 |
| Mermaid, не image_gen | 8, 12 |
| report.md конвенции | 9, 10 |
| клоны figure/table/абзац, подписи {n} | 10 |
| стоп без прототипа рисунка/таблицы | 10 |
| screenshot без ok | 10 |
| render Word затем LO, skip теста | 11 |
| визуальный QA, 3 цикла, /report | 12 |
| не Composio Word | 1 AGENTS, 12 |
| тесты 1–5 из спеки | 4, 5, 6, 10, 11 |
