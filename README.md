# Фабрика учебных отчётов

Локальный конвейер для Grok: PDF задания → Markdown с картинками → выполнение работы со скриншотами → `.docx` **правкой копии эталона** (титульник и стили не собираются с нуля).

## Установка

```
pip install -r requirements.txt
```

Скопируйте `student.yaml.example` в `student.yaml` и заполните ФИО/группу.

Для проверки вёрстки нужен Microsoft Word (предпочтительно на Windows) или LibreOffice. Для диаграмм — Node.js (`npx @mermaid-js/mermaid-cli`).

## Как пользоваться

1. Положите эталон другой работы: `subjects/<предмет>/sample.docx`.
2. В чате Grok из корня репозитория: `/report` и PDF задания (для нового предмета Grok предложит `title-fields.yaml` — подтвердите).
3. Готовый файл: `subjects/<предмет>/lab-NN-<slug>/report.docx`.

Не редактируйте эталон. Сломанную копию пересоберите: `python tools/new_work.py --subject S --slug SLUG --number N --reset-docx`.

## Скрипты

| Скрипт | Назначение |
|---|---|
| `tools/pdf_to_md.py` | PDF → `assignment.md` + `images/` |
| `tools/new_work.py` | папка работы, копия эталона |
| `tools/patch_title.py` | замена ФИО/группы/темы на титульнике |
| `tools/replace_body.py` | тело отчёта клонами блоков эталона |
| `tools/screenshot.py` | PNG + `manifest.json` |
| `tools/diagram.py` | Mermaid → PNG |
| `tools/render_docx.py` | страницы отчёта → PNG |

Спека: `docs/superpowers/specs/2026-09-11-report-factory-design.md`.
