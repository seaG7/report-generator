from tools.parse_report import parse_report_md

MD = (
    "# Цель работы\n\n"
    "Текст цели.\n\n"
    "![Установка зависимостей](screenshots/01-install.png)\n\n"
    "## Ход\n\n"
    "Таблица: Результаты измерений\n\n"
    "| Параметр | Значение |\n"
    "| --- | --- |\n"
    "| t | 12 |\n\n"
    "```python\n"
    'print("hello")\n'
    "```\n"
)


def test_parse_report_md_kinds():
    blocks = parse_report_md(MD)
    kinds = [b.kind for b in blocks]
    assert kinds == ["heading", "paragraph", "figure", "heading", "table", "code"]
    assert blocks[0].level == 1 and blocks[0].text == "Цель работы"
    assert blocks[2].caption == "Установка зависимостей"
    assert blocks[2].path == "screenshots/01-install.png"
    assert blocks[4].caption == "Результаты измерений"
    assert blocks[4].headers == ["Параметр", "Значение"]
    assert blocks[4].rows == [["t", "12"]]
    assert blocks[5].lang == "python" and "print" in blocks[5].text
