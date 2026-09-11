from pathlib import Path
import pytest
from tests.fixtures_util import make_sample_docx
from tools.render_docx import RenderUnavailable, render_docx, soffice_available, word_available


def test_render_unavailable_message():
    err = RenderUnavailable("Нет Word и нет LibreOffice")
    assert "LibreOffice" in str(err)


@pytest.mark.skipif(
    not word_available() and not soffice_available(),
    reason="Нет Word и нет LibreOffice",
)
def test_render_docx_writes_nonempty_png(tmp_path: Path):
    docx = tmp_path / "s.docx"
    make_sample_docx(docx)
    out = tmp_path / "preview"
    pages = render_docx(docx, out)
    assert pages, "ожидалась хотя бы одна страница"
    assert pages[0].name == "page-01.png"
    assert pages[0].stat().st_size > 0
