---
name: report
description: Run the full lab-report factory from assignment PDF plus subject sample.docx to a verified report.docx. Use when the user wants a lab report, /report, лабораторная, сделай отчёт, or attaches an assignment PDF with a formatted sample.
---

# Фабрика отчёта

Follow `docs/superpowers/specs/2026-09-11-report-factory-design.md`. Commands from repo root. Never write `subjects/<subject>/sample.docx`. Never build Word via Composio or `Document()` from scratch.

## 0. Старт

1. If `student.yaml` is missing, stop: copy `student.yaml.example` → `student.yaml` and ask the user to fill it.
2. Resolve subject (argument, cwd under `subjects/<slug>/`, or one question).
3. If `subjects/<subject>/sample.docx` is missing, stop: need a formatted sample of another work in this discipline.
4. If `title-fields.yaml` is missing or `confirmed` is not true — onboarding only, do not execute the lab:
   - Render the sample for reading only: `python tools/render_docx.py "subjects/<subject>/sample.docx" --out-dir <tmp>/preview`
   - Read `page-01.png` and title text. Propose `find` strings, `body_start`, `figure_caption`, `table_caption`.
   - Wait for user confirmation, then write `title-fields.yaml` with `confirmed: true`.
5. `python tools/new_work.py --subject S --slug SLUG [--pdf PDF] [--number N]`

Rebuild a ruined copy: `python tools/new_work.py --subject S --slug SLUG --number N --reset-docx` (does not wipe `work/` or `report.md`).

## 1. Ingest

`python tools/pdf_to_md.py "<assignment.pdf>" --out "<work-dir>"`

Read `assignment.md`. Do not invent missing PDF figures.

## 2. Execute

Do the assignment in `<work-dir>/work/`.

Screenshots:

```
python tools/screenshot.py --from-file IMG --out-dir "<work>/screenshots" --caption "..." --source browser|terminal|ide --slug slug
python tools/screenshot.py --capture --out-dir "<work>/screenshots" --caption "..." --source ide --slug slug
```

Diagrams: write `work/diagrams/<name>.mmd`, then

```
python tools/diagram.py "<work>/work/diagrams/<name>.mmd"
python tools/screenshot.py --from-file PNG --out-dir "<work>/screenshots" --caption "..." --source diagram --slug slug
```

Do not use image generation for flowcharts, sequences, or architecture if Mermaid can draw them.

Open every PNG, check it against the assignment, then:

```
python tools/screenshot.py --mark-ok FILE.png --out-dir "<work>/screenshots"
```

Never set `ok` without looking. Write `work.yaml` and `report.md` (same section structure as the sample; `![caption](screenshots/...)`; `Таблица: ...` before a GFM table).

## 3–4. Compose and verify

Follow the `build-report` skill. Max 3 inspect/fix loops. Ready artifact: `<work>/report.docx` after a passing page inspect.
