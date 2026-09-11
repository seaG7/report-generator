from pathlib import Path
import re
from tests.fixtures_util import make_two_image_pdf
from tools.pdf_to_md import convert_pdf


def test_images_appear_in_document_order(tmp_path: Path):
    pdf = tmp_path / "a.pdf"
    make_two_image_pdf(pdf)
    out = tmp_path / "out"
    result = convert_pdf(pdf, out)
    md = result.md_path.read_text(encoding="utf-8")
    assert result.md_path == out / "assignment.md"
    assert (out / "images").is_dir()
    assert len(result.image_paths) >= 2
    pos_before = md.find("BEFORE")
    pos_middle = md.find("MIDDLE")
    pos_after = md.find("AFTER")
    imgs = [m.start() for m in re.finditer(r"!\[[^\]]*\]\([^)]+\)", md)]
    assert pos_before != -1 and pos_middle != -1 and pos_after != -1
    assert len(imgs) >= 2
    assert pos_before < imgs[0] < pos_middle < imgs[1] < pos_after
    for p in result.image_paths:
        assert p.is_file() and p.stat().st_size > 0
