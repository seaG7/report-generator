---
name: report
description: Run the full lab-report factory. Drop PDFs into subjects/<slug>/inbox/ then /report. Inbox may mix assignment, lectures, and other materials. Use when the user wants a lab report, лабораторная, сделай отчёт, or has PDFs in a subject inbox.
---

# Фабрика отчёта

Команды из корня репо. Эталон `subjects/<subject>/sample.docx` не писать. Word не собирать с нуля и не через Composio.

Пользователь **не прикладывает PDF в чат**. В `subjects/<предмет>/inbox/` (или рядом с `sample.docx`) может лежать **несколько** PDF: задание, лекции, слайды, методичка.

## 0. Превью всех PDF — не создавать лабу сразу

Если нет `student.yaml` — скопировать из `student.yaml.example` и остановиться, пока не заполнят ФИО.

```
python tools/pickup.py --preview --json
```

- код 0 — JSON: `[{subject, pdf, md}, ...]` на **каждый** PDF, файлы на месте не трогаются
- код 2 — PDF нет: положить в `subjects/<предмет>/inbox/` (`python tools/new_subject.py --slug <slug>`)
- код 1 — ошибка конвертации

Прочитать **каждый** `md` целиком (и картинки в той же папке превью). Классифицировать по содержанию, не по имени файла:

- `assignment` — лабораторная/практика/контрольная: «выполнить», варианты, цели работы, ход работы, задания по пунктам
- `lecture` — лекция, конспект, теория, слайды курса без требования сдать работу
- `other` — силлабус, методичка, правила оформления, чужой пример не как эталон Word

Несколько PDF задания (часть 1 и 2, два файла одного варианта) — **одна** лаба, все в `--assignment`. Две независимые лабораторные — два `--claim`. Лекции не превращать в отдельные lab-папки.

Нет `title-fields.yaml` или `confirmed` не true — онбординг эталона, лабу не выполнять.

Затем:

```
python tools/pickup.py --claim --subject S --assignment "<pdf-задания>" [--assignment "<часть-2.pdf>"] [--lecture "<лекция.pdf>"] [--other "<методичка.pdf>"]
```

Повторить `--claim`, если в inbox два разных задания. Лекции окажутся в `subjects/<S>/lectures/` и в `<work>/lectures/`.

## 1. Ingest

Для каждого PDF задания в папке работы (`assignment.pdf`, `assignment-02-*.pdf`):

```
python tools/pdf_to_md.py "<work>/assignment.pdf" --out "<work>"
python tools/checklist.py "<work>/assignment.md" --out "<work>/work/checklist.md"
```

Доп. части задания — тоже в MD (`assignment-02.md`) и **добавить** пункты в тот же чеклист, ничего не отбрасывая.

Каждую лекцию:

```
python tools/pdf_to_md.py "<work>/lectures/<file>.pdf" --out "<work>/lectures/<stem>"
```

Читать лекции как контекст. Чеклист и отчёт — только по **заданию**. Открыть все картинки задания.

## 2. Выполнить всё задание

Закрыть каждый пункт чеклиста. Нельзя сдать отчёт, пока в `checklist.md` есть `- [ ]`.

Контекст: текст, списки, таблицы, подписи, изображения задания; лекции — опора, не замена пунктов.

Инструменты, если нужного софта нет:

1. MCP: `search_tool`.
2. Composio: `COMPOSIO_SEARCH_TOOLS` / `COMPOSIO_MANAGE_CONNECTIONS`.
3. GitHub/CLI (`web_search`): Newman, httpie, k6 и т.п. Поставить и использовать.
4. Не писать «не могу, нет Postman/MCP».

Скриншоты: `screenshot.py` / `diagram.py`, осмотр PNG, `--mark-ok`. Mermaid вместо image generation, если хватает.

`work.yaml` + `report.md` по структуре эталона.

## 3–4. Сборка

Скилл `build-report`. Максимум 3 цикла осмотра preview. Готово: `<work>/report.docx`.

Сломанная копия: `python tools/new_work.py --subject S --slug SLUG --number N --reset-docx`.
