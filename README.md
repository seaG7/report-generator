# Фабрика учебных отчётов

Кидаете PDF в папку предмета, пишете `/report`. Сортировать файлы не нужно.

## Один раз

```
pip install -r requirements.txt
copy student.yaml.example student.yaml
```

Заполните ФИО и группу в `student.yaml`.

```
python tools/new_subject.py --slug os
```

Эталон другой работы: `subjects/os/sample.docx`.

Нужны Word или LibreOffice (проверка вёрстки). Для диаграмм — Node.js.

## Каждая работа

1. Все PDF (задание, лекции, что угодно) → `subjects/os/`
2. В Grok: `/report`

Grok сам читает файлы, отличает задание от лекций и делает отчёт. Готово: `subjects/os/lab-NN-.../report.docx`.

Первый раз по предмету подтвердите поля титульника.

## Скрипты

| Скрипт | Назначение |
|---|---|
| `tools/new_subject.py` | папка предмета |
| `tools/pickup.py` | превью PDF в папке предмета; забрать в lab |
| `tools/pdf_to_md.py` | PDF → markdown + картинки |
| `tools/checklist.py` | чеклист пунктов задания |
| `tools/new_work.py` | папка работы, копия эталона |
| `tools/patch_title.py` | ФИО/группа/тема на титульнике |
| `tools/replace_body.py` | тело отчёта клонами блоков |
| `tools/screenshot.py` | PNG + манифест |
| `tools/diagram.py` | Mermaid → PNG |
| `tools/render_docx.py` | страницы → PNG |

Спека: `docs/superpowers/specs/2026-09-11-report-factory-design.md`.
