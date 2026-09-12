---
name: report
description: Run the lab-report factory. User dumps PDFs into subjects/<slug>/ (assignment mixed with lectures) then /report. Use when the user wants a lab report, лабораторная, сделай отчёт, or has PDFs in a subject folder.
---

# Фабрика отчёта

Команды из корня репо. Эталон `sample.docx` не писать. Word не собирать с нуля и не через Composio.

Пользователь **не сортирует файлы и не прикладывает PDF в чат**. Всё лежит прямо в `subjects/<предмет>/`: задание, лекции, слайды — одним ворохом.

Не создавать inbox/lectures/materials. Не просить разложить PDF по папкам.

## 0. Превью

Нет `student.yaml` — скопировать из `student.yaml.example`, остановиться пока не заполнят ФИО.

```
python tools/pickup.py --preview --json
```

- код 0 — JSON `[{subject, pdf, md}, ...]` на каждый PDF **в корне предмета** (не внутри `lab-*`)
- код 2 — PDF нет: «положите файлы в subjects/<предмет>/»
- код 1 — ошибка

Прочитать каждый `md` целиком и картинки. По содержанию, не по имени:

- задание: лабораторная/практика, «выполнить», варианты, цель/ход работы
- контекст: лекция, конспект, слайды, методичка — пригодится при выполнении, но это не сама работа

Несколько файлов одного задания (часть 1 и 2) — одна лаба, все в `--assignment`. Две разные лабораторные — два `--claim`. Контекст не делать отдельной lab-папкой.

Нет `title-fields.yaml` / `confirmed` — онбординг эталона, работу не начинать.

```
python tools/pickup.py --claim --subject S --assignment "<задание.pdf>" [--assignment "<часть-2.pdf>"] [--context "<лекция.pdf>"] [--context "<ещё.pdf>"]
```

Все эти PDF окажутся **плоско** в папке работы (без подпапок типа lectures/). Корень предмета снова пустой от PDF — следующую работу снова кидают туда же.

## 1. Ingest

```
python tools/pdf_to_md.py "<work>/assignment.pdf" --out "<work>"
python tools/checklist.py "<work>/assignment.md" --out "<work>/work/checklist.md"
```

Доп. части задания — в MD и в тот же чеклист. Каждый контекстный PDF:

```
python tools/pdf_to_md.py "<work>/<файл>.pdf" --out "<work>/.context/<stem>"
```

Чеклист и отчёт — по заданию. Контекст читать и использовать. Картинки задания открыть все.

## 2. Выполнить всё задание

Пока в `checklist.md` есть `- [ ]` — отчёт не готов.

Нет инструмента:

1. MCP `search_tool`
2. Composio connect
3. GitHub/CLI (`web_search`)
4. Не писать «не могу, нет Postman»

Скриншоты: `screenshot.py` / `diagram.py`, осмотр, `--mark-ok`. Mermaid вместо image generation, если хватает.

`work.yaml` + `report.md` по эталону.

## 3–4. Сборка

Скилл `build-report`. Максимум 3 цикла preview. Готово: `<work>/report.docx`.

Сброс копии: `python tools/new_work.py --subject S --slug SLUG --number N --reset-docx`.
