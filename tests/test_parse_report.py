from tools.parse_report import parse_report_md

MD = (
    "# Цель работы\n\n"
    "Текст цели.\n\n"
    "– пункт задания\n\n"
    "![Установка зависимостей](screenshots/01-install.png)\n\n"
    "## Ход\n\n"
    "Таблица: Результаты измерений\n\n"
    "| Параметр | Значение |\n"
    "| --- | --- |\n"
    "| t | 12 |\n\n"
    "Листинг: Создание базы\n\n"
    "```python\n"
    'print("hello")\n'
    "```\n"
)


def test_parse_report_md_kinds():
    blocks = parse_report_md(MD)
    kinds = [b.kind for b in blocks]
    assert kinds == ["heading", "paragraph", "dash", "figure", "heading", "table", "code"]
    assert blocks[0].level == 1 and blocks[0].text == "Цель работы"
    assert blocks[2].kind == "dash" and blocks[2].text.startswith("– ")
    assert blocks[3].caption == "Установка зависимостей"
    assert blocks[3].path == "screenshots/01-install.png"
    assert blocks[5].caption == "Результаты измерений"
    assert blocks[5].headers == ["Параметр", "Значение"]
    assert blocks[5].rows == [["t", "12"]]
    assert blocks[6].lang == "python" and "print" in blocks[6].text
    assert blocks[6].caption == "Создание базы"


def test_split_row_keeps_escaped_pipe():
    from tools.parse_report import parse_report_md

    blocks = parse_report_md(
        "# Цель\n\nТаблица: T\n\n"
        "| Сайт | Title |\n"
        "| --- | --- |\n"
        "| X | foo \\| bar |\n"
    )
    tbl = next(b for b in blocks if b.kind == "table")
    assert tbl.rows == [["X", "foo | bar"]]
