from tools.checklist import extract_checklist


def test_extract_checklist_covers_headings_items_and_images():
    md = (
        "# Цель\n\n"
        "Сделать API.\n\n"
        "1. Создать коллекцию в Postman\n"
        "2. Проверить GET /health\n\n"
        "- включить CORS\n\n"
        "![Схема](images/scheme.png)\n\n"
        "## Выводы\n"
    )
    items = extract_checklist(md)
    texts = [i.text for i in items]
    assert "Цель" in texts
    assert "Создать коллекцию в Postman" in texts
    assert "Проверить GET /health" in texts
    assert "включить CORS" in texts
    assert any(i.kind == "image" and "scheme.png" in i.ref for i in items)
    assert "Выводы" in texts
