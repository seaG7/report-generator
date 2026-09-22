from __future__ import annotations

from dataclasses import dataclass
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
    listing_caption: str = "Листинг {n} — {caption}"


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
        listing_caption=data.get("listing_caption") or "Листинг {n} — {caption}",
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
