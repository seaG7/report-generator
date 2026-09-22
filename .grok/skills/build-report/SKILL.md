---
name: build-report
description: Patch title fields and replace body of a copied sample.docx, then render pages for visual QA. Use when composing the Word report after work is done, or when the user runs /build-report.
---

# Сборка report.docx

Preconditions: `title-fields.yaml` has `confirmed: true`; `work.yaml` exists; `report.md` exists; `work/checklist.md` has no unchecked `- [ ]` items; every `screenshots/` image referenced in `report.md` has `ok: true` in `manifest.json`. Do not compose if any assignment checklist item is still open.

```
python tools/patch_title.py --docx "<work>/report.docx" --title-fields "subjects/<subject>/title-fields.yaml" --student student.yaml --work "<work>/work.yaml"
python tools/replace_body.py --docx "<work>/report.docx" --md "<work>/report.md" --title-fields "subjects/<subject>/title-fields.yaml" --work-dir "<work>"
python tools/fit_figures.py --docx "<work>/report.docx"
python tools/continue_tables.py --docx "<work>/report.docx"
python tools/render_docx.py "<work>/report.docx" --out-dir "<work>/preview"
```

If `<work>/student.yaml` exists, `patch_title` uses it instead of the repo-root `student.yaml` (second author of the same practice).

Never write `sample.docx`. Never create a Word document from scratch. After render, open every `preview/page-*.png` and inspect against the sample and these factory blocks:

- title intact; body first-line indent and justify
- figure captions 12 pt bold center; images full text width; no large empty region above a figure that still fits if shrunk (`fit_figures.py` after layout)
- assignment lists as «– » paragraphs (not Word bullets)
- listings: italic left caption, Courier New 10 pt in a single frame
- never stack figure/figure, figure/table or table/table without a text paragraph between them
- table rows have the same cell count as the header; leftover cells mean a raw `|` in a cell
- «Продолжение Таблицы N» only on the page where Word actually continued the table, after layout (`continue_tables.py`); do not pre-split rows in replace_body; caption is a paragraph above the continued chunk, never a table cell
- every numbered H1 (1 …, 2 …) starts a new page; ЦЕЛЬ/ЗАДАНИЕ stay with the previous section
- ЗАКЛЮЧЕНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ, ПРИЛОЖЕНИЯ start a new page; «Приложение А/Б/…» follows ПРИЛОЖЕНИЯ on the same page

On failure, fix the copy and re-render. After 3 failed inspect loops, stop and list defects. Do not declare the report ready without a passing inspect.
