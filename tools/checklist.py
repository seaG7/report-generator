from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ChecklistItem:
    kind: str  # heading | item | image
    text: str
    ref: str = ""


_HEADING = re.compile(r"^(#{1,3})\s+(.+?)\s*$")
_NUMBERED = re.compile(r"^\s*\d+[.)]\s+(.+?)\s*$")
_BULLET = re.compile(r"^\s*[-*+]\s+(.+?)\s*$")
_IMAGE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")


def extract_checklist(md: str) -> list[ChecklistItem]:
    items: list[ChecklistItem] = []
    for line in md.replace("\r\n", "\n").split("\n"):
        m_h = _HEADING.match(line)
        if m_h:
            items.append(ChecklistItem(kind="heading", text=m_h.group(2).strip()))
            continue
        m_n = _NUMBERED.match(line)
        if m_n:
            items.append(ChecklistItem(kind="item", text=m_n.group(1).strip()))
            continue
        m_b = _BULLET.match(line)
        if m_b:
            items.append(ChecklistItem(kind="item", text=m_b.group(1).strip()))
            continue
        for alt, path in _IMAGE.findall(line):
            caption = alt.strip() or Path(path).name
            items.append(ChecklistItem(kind="image", text=caption, ref=path.strip()))
    return items


def write_checklist(path: Path, items: list[ChecklistItem]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = ["# Чеклист задания", ""]
    for item in items:
        extra = f" ({item.ref})" if item.ref else ""
        lines.append(f"- [ ] [{item.kind}] {item.text}{extra}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("md")
    p.add_argument("--out", required=True)
    args = p.parse_args()
    md_path = Path(args.md)
    if not md_path.is_file():
        print(f"Нет файла: {md_path}", file=sys.stderr)
        return 1
    items = extract_checklist(md_path.read_text(encoding="utf-8"))
    write_checklist(Path(args.out), items)
    print(args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
