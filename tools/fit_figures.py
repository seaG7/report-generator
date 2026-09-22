"""Shrink figures that Word pushed to the next page so leftover space is filled."""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from docx import Document
from lxml import etree

from tools.insert_figure import scale_drawing_to_height
from tools.paths import assert_not_sample
from tools.render_docx import word_available


class FitError(Exception):
    pass


_CAP = re.compile(r"^Рисунок\s+(\d+)\b")
WD_ACTIVE_END_PAGE = 3
WD_VERTICAL_POS_PAGE = 6
EMU_PER_PT = 12700
MIN_FIGURE_HEIGHT_EMU = 800_000  # ~2.2 cm
RESERVE_PT = 28  # caption line + spacing around the figure block
SAFETY_PT = 6


def _para_bottom_pt(para) -> tuple[int, float]:
    rng = para.Range.Duplicate
    start, end = int(rng.Start), int(rng.End)
    if end > start:
        rng.SetRange(end - 1, end)
    page = int(rng.Information(WD_ACTIVE_END_PAGE))
    y = float(rng.Information(WD_VERTICAL_POS_PAGE))
    ls = float(para.Range.ParagraphFormat.LineSpacing)
    if ls <= 6:
        size = float(para.Range.Font.Size) or 12.0
        ls = max(ls, 1.0) * size
    return page, y + ls


def _collect_fits_word_impl(docx_path: Path) -> list[dict]:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    fits: list[dict] = []
    try:
        doc = word.Documents.Open(str(docx_path.resolve()), False, True, False)
        try:
            doc.Repaginate()
        except Exception:
            pass
        page_h = float(doc.PageSetup.PageHeight)
        bot = float(doc.PageSetup.BottomMargin)
        count = int(doc.Paragraphs.Count)
        for i in range(1, count):
            para = doc.Paragraphs(i)
            try:
                if int(para.Range.InlineShapes.Count) < 1:
                    continue
            except Exception:
                continue
            nxt = doc.Paragraphs(i + 1)
            cap = (nxt.Range.Text or "").replace("\r", " ").replace("\x07", " ").strip()
            m = _CAP.match(cap)
            if not m:
                continue
            if i <= 1:
                continue
            prev = doc.Paragraphs(i - 1)
            try:
                if bool(prev.Range.Information(12)):
                    continue
                fig_page = int(para.Range.Information(WD_ACTIVE_END_PAGE))
                prev_page, prev_bottom = _para_bottom_pt(prev)
            except Exception:
                continue
            if fig_page <= prev_page:
                continue
            try:
                cur_h_pt = float(para.Range.InlineShapes(1).Height)
            except Exception:
                continue
            leftover_pt = page_h - bot - prev_bottom
            target_pt = leftover_pt - RESERVE_PT - SAFETY_PT
            target_cy = int(target_pt * EMU_PER_PT)
            cur_cy = int(cur_h_pt * EMU_PER_PT)
            if target_cy < MIN_FIGURE_HEIGHT_EMU:
                continue
            if target_cy >= cur_cy - EMU_PER_PT * 2:
                continue
            fits.append({"n": int(m.group(1)), "cy": target_cy})
            break
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
    return fits


def _apply_fits(docx_path: Path, fits: list[dict]) -> int:
    if not fits:
        return 0
    document = Document(str(docx_path))
    applied = 0
    wanted = {int(s["n"]): int(s["cy"]) for s in fits}
    body = list(document.element.body)
    for idx, child in enumerate(body):
        if etree.QName(child).localname != "p":
            continue
        text = (child.xpath("string()") or "").strip()
        m = _CAP.match(text)
        if not m:
            continue
        n = int(m.group(1))
        if n not in wanted:
            continue
        prev = body[idx - 1] if idx else None
        if prev is None or etree.QName(prev).localname != "p":
            continue
        if scale_drawing_to_height(prev, wanted[n]):
            applied += 1
    if applied:
        document.save(str(docx_path))
    return applied


def _collect_fits(docx_path: Path) -> list[dict]:
    env = os.environ.copy()
    env["_REPORT_WORD_WORKER"] = "1"
    proc = subprocess.run(
        [sys.executable, "-m", "tools.fit_figures", "--docx", str(docx_path), "--collect"],
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.returncode != 0:
        raise FitError((proc.stderr or proc.stdout or "Word не смог разметить рисунки").strip())
    raw = (proc.stdout or "").strip()
    try:
        return json.loads(raw) if raw else []
    except json.JSONDecodeError as exc:
        raise FitError(f"Не разобрать разметку рисунков: {exc}") from exc


def fit_figures(docx_path: Path) -> int:
    docx_path = Path(docx_path)
    assert_not_sample(docx_path)
    if not word_available():
        return 0
    total = 0
    for _ in range(8):
        fits = _collect_fits(docx_path)
        if not fits:
            break
        applied = _apply_fits(docx_path, fits[:1])
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
            print(json.dumps(_collect_fits_word_impl(Path(args.docx)), ensure_ascii=False))
            return 0
        n = fit_figures(Path(args.docx))
        print(f"подогнано рисунков: {n}")
    except FitError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
