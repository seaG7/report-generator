from pathlib import Path
from PIL import Image
import json
import pytest
from tools.screenshot import ManifestError, mark_ok, register_image


def test_register_image_writes_manifest_ok_false(tmp_path: Path):
    src = tmp_path / "a.png"
    Image.new("RGB", (10, 10), (1, 2, 3)).save(src)
    out = tmp_path / "screenshots"
    dest = register_image(src, out, "Установка зависимостей", "terminal", slug="install")
    assert dest.name == "01-install.png"
    assert dest.is_file()
    data = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
    assert data == [{
        "file": "01-install.png",
        "caption": "Установка зависимостей",
        "source": "terminal",
        "ok": False,
    }]
    mark_ok(out, "01-install.png")
    data = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
    assert data[0]["ok"] is True


def test_second_image_increments_index(tmp_path: Path):
    src = tmp_path / "a.png"
    Image.new("RGB", (10, 10)).save(src)
    out = tmp_path / "screenshots"
    register_image(src, out, "a", "browser", slug="a")
    dest = register_image(src, out, "b", "browser", slug="b")
    assert dest.name == "02-b.png"


def test_mark_ok_missing_raises(tmp_path: Path):
    out = tmp_path / "screenshots"
    out.mkdir()
    (out / "manifest.json").write_text("[]", encoding="utf-8")
    with pytest.raises(ManifestError):
        mark_ok(out, "nope.png")
