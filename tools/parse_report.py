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
_DASH = re.compile(r"^[–—\-]\s+(.+?)\s*$")


def _split_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _is_separator(cells: list[str]) -> bool:
    if not cells:
        return False
    return all(set(c.replace(":", "").replace("-", "").strip()) == set() and "-" in c for c in cells)


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
        m_dash = _DASH.match(line.strip())
        if m_dash:
            flush_para()
            blocks.append(Block(kind="dash", text="– " + m_dash.group(1)))
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
            if len(rows) >= 2 and _is_separator(rows[1]):
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
