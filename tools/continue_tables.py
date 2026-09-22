"""Split tables only where Word actually broke them across pages; insert «Продолжение Таблицы N»."""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from tools.paths import assert_not_sample
from tools.render_docx import word_available
from tools.replace_body import _nil_bottom_border


class ContinueError(Exception):
    pass


_CAP_N = re.compile(r"(?:Таблица|Таблицы)\s+(\d+)")
WD_ACTIVE_END_PAGE = 3


def _caption_n(text: str) -> int | None:
    m = _CAP_N.search(text or "")
    return int(m.group(1)) if m else None


def _collect_splits_word_impl(docx_path: Path) -> list[dict]:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    splits: list[dict] = []
    try:
        doc = word.Documents.Open(str(docx_path.resolve()), False, True, False)
        try:
            doc.Repaginate()
        except Exception:
            pass
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            if tbl.Columns.Count <= 1 or tbl.Rows.Count < 3:
                continue
            pages = [
                int(tbl.Rows(ri).Range.Information(WD_ACTIVE_END_PAGE))
                for ri in range(1, tbl.Rows.Count + 1)
            ]
            first = pages[0]
            split_at = next((i + 1 for i, p in enumerate(pages) if p != first), None)
            # Need header + at least one data row on the first page, otherwise
            # the first chunk would be a header-only stub.
            if split_at is None or split_at <= 2:
                continue
            pos = int(tbl.Range.Start)
            chunk = doc.Range(max(0, pos - 500), pos).Text
            n = _caption_n(chunk) or ti
            splits.append({"ti": ti - 1, "split_row": split_at - 1, "n": n})
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    return splits


def _continuation_paragraph(n: int):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(OxmlElement("w:keepNext"))
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "left")
    pPr.append(jc)
    p.append(pPr)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rPr.append(OxmlElement("w:i"))
    rPr.append(OxmlElement("w:iCs"))
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = f"Продолжение Таблицы {n}"
    r.append(t)
    p.append(r)
    return p


def _split_tbl_at(tbl, split_row: int, n: int) -> None:
    trs = tbl.findall(qn("w:tr"))
    if split_row <= 1 or split_row >= len(trs):
        return
    new_tbl = deepcopy(tbl)
    new_trs = new_tbl.findall(qn("w:tr"))
    for tr in list(new_trs[1:split_row]):
        new_tbl.remove(tr)
    for tr in list(trs[split_row:]):
        tbl.remove(tr)
    _nil_bottom_border(tbl)
    cap = _continuation_paragraph(n)
    tbl.addnext(new_tbl)
    tbl.addnext(cap)


def _apply_splits(docx_path: Path, splits: list[dict]) -> int:
    if not splits:
        return 0
    document = Document(str(docx_path))
    tables = list(document.tables)
    applied = 0
    for spec in sorted(splits, key=lambda s: s["ti"], reverse=True):
        ti = spec["ti"]
        if ti < 0 or ti >= len(tables):
            continue
        _split_tbl_at(tables[ti]._tbl, spec["split_row"], spec["n"])
        applied += 1
    document.save(str(docx_path))
    return applied


def _collect_splits(docx_path: Path) -> list[dict]:
    env = os.environ.copy()
    env["_REPORT_WORD_WORKER"] = "1"
    proc = subprocess.run(
        [sys.executable, "-m", "tools.continue_tables", "--docx", str(docx_path), "--collect"],
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.returncode != 0:
        raise ContinueError((proc.stderr or proc.stdout or "Word не смог разметить продолжения таблиц").strip())
    raw = (proc.stdout or "").strip()
    try:
        return json.loads(raw) if raw else []
    except json.JSONDecodeError as exc:
        raise ContinueError(f"Не разобрать разметку таблиц: {exc}") from exc


def continue_tables(docx_path: Path) -> int:
    docx_path = Path(docx_path)
    assert_not_sample(docx_path)
    if not word_available():
        return 0
    total = 0
    for _ in range(6):
        splits = _collect_splits(docx_path)
        if not splits:
            break
        applied = _apply_splits(docx_path, splits)
        if applied <= 0:
            break
        total += applied
    return total


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--docx", required=True)
    p.add_argument("--collect", action="store_true")
    args = p.parse_args()
    try:
        if args.collect:
            print(json.dumps(_collect_splits_word_impl(Path(args.docx)), ensure_ascii=False))
            return 0
        n = continue_tables(Path(args.docx))
        print(f"продолжений: {n}")
    except ContinueError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
