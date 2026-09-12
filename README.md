# Фабрика учебных отчётов

Кидаете PDF в папку предмета, в чате пишете `/report` — Grok делает работу и собирает Word копией эталона.

## Один раз

```
pip install -r requirements.txt
copy student.yaml.example student.yaml
```

Заполните ФИО и группу в `student.yaml`.

Для каждого предмета:

```
python tools/new_subject.py --slug os
```

Положите эталон другой работы: `subjects/os/sample.docx`.

Для проверки вёрстки нужен Word или LibreOffice. Для диаграмм — Node.js.

## Каждая лабораторная

1. Все PDF по работе → `subjects/os/inbox/` (задание, лекции, методички — можно вместе).
2. В Grok из корня репо: `/report`.

Grok читает каждый PDF и сам отделяет задание от лекций. В новую лабу попадает только задание; лекции складываются в `subjects/os/lectures/` и копируются в папку работы.

Готово: `subjects/os/lab-NN-.../report.docx`.

Первый раз по предмету Grok покажет поля титульника — подтвердите. Дальше не спрашивает.

## Скрипты

| Скрипт | Назначение |
|---|---|
| `tools/new_subject.py` | папка предмета + `inbox/` |
| `tools/pickup.py` | превью inbox PDF; после разбора — claim задания/лекций |
| `tools/pdf_to_md.py` | PDF → `assignment.md` + картинки |
| `tools/checklist.py` | чеклист пунктов и рисунков задания |
| `tools/new_work.py` | папка работы, копия эталона |
| `tools/patch_title.py` | ФИО/группа/тема на титульнике |
| `tools/replace_body.py` | тело отчёта клонами блоков |
| `tools/screenshot.py` | PNG + манифест |
| `tools/diagram.py` | Mermaid → PNG |
| `tools/render_docx.py` | страницы → PNG |

Спека: `docs/superpowers/specs/2026-09-11-report-factory-design.md`.
