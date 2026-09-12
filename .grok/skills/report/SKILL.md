---
name: report
description: Run the full lab-report factory. Drop the assignment PDF into subjects/<slug>/inbox/ then /report. Use when the user wants a lab report, лабораторная, сделай отчёт, or has a PDF in a subject inbox.
---

# Фабрика отчёта

Команды из корня репо. Эталон `subjects/<subject>/sample.docx` не писать. Word не собирать с нуля и не через Composio.

Пользователь **не прикладывает PDF в чат**. Файл лежит в `subjects/<предмет>/inbox/` или рядом с `sample.docx`.

## 0. Подхватить PDF

Если нет `student.yaml` — скопировать из `student.yaml.example` и остановиться, пока не заполнят ФИО.

```
python tools/pickup.py
```

- код 0 — по одной строке путь к каждой новой `lab-NN-...`
- код 2 — PDF нет: сказать положить в `subjects/<предмет>/inbox/` (создать предмет: `python tools/new_subject.py --slug <slug>`)
- код 1 — нет эталона: нужен `subjects/<предмет>/sample.docx`

Нет `title-fields.yaml` или `confirmed` не true — онбординг: рендер эталона только на чтение, предложить поля, дождаться подтверждения, записать `confirmed: true`. Лабу не выполнять до этого.

Дальше для **каждой** папки из pickup.

## 1. Ingest

```
python tools/pdf_to_md.py "<work>/assignment.pdf" --out "<work>"
python tools/checklist.py "<work>/assignment.md" --out "<work>/work/checklist.md"
```

Прочитать `assignment.md` **целиком**. Открыть каждую картинку из `images/`. Чеклист — обязательные пункты; не вычёркивать то, чего не сделали.

Не выдумывать рисунки, которых нет в извлечении.

## 2. Выполнить всё задание

Закрыть каждый пункт чеклиста. Нельзя сдать отчёт, пока в `checklist.md` есть `- [ ]`.

Контекст: текст, нумерованные списки, буллеты, таблицы, подписи и изображения из PDF. Если картинка в задании — она часть требования.

Инструменты, если в задании нужен софт, которого нет под рукой:

1. Сначала MCP: `search_tool` по имени (Postman, браузер, git, …) и вызвать найденное.
2. Нет коннекта — Composio: `COMPOSIO_SEARCH_TOOLS`, при необходимости `COMPOSIO_MANAGE_CONNECTIONS` (дать пользователю ссылку, дождаться ACTIVE).
3. Нет MCP — найти готовое на GitHub (`web_search` / GitHub): CLI вроде Newman для Postman-коллекций, httpie, k6, и т.п. Поставить и использовать.
4. Не писать «не могу, нет Postman/MCP». Либо подключить, либо эквивалент.

Скриншоты и диаграммы — как раньше (`screenshot.py`, `diagram.py`, осмотр PNG, `--mark-ok`). Mermaid вместо image generation, если хватает.

`work.yaml` + `report.md` по структуре эталона. В отчёт входят все пункты чеклиста.

## 3–4. Сборка

Скилл `build-report`. Максимум 3 цикла осмотра preview. Готово: `<work>/report.docx`.

Сломанная копия: `python tools/new_work.py --subject S --slug SLUG --number N --reset-docx`.
