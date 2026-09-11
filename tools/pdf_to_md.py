from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import pymupdf
import pymupdf4llm


class ConvertError(Exception):
    pass


@dataclass
class ConvertResult:
    md_path: Path
    image_paths: list[Path]
    warning: str | None
    empty_text: bool


_IMG_MD = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")


def _rewrite_image_links(md: str, images_dir: Path) -> str:
    def repl(match: re.Match[str]) -> str:
        alt, raw = match.group(1), match.group(2)
        name = Path(raw.replace("\\", "/")).name
        return f"![{alt}](images/{name})"

    return _IMG_MD.sub(repl, md)


def _image_paths_in_md_order(md: str, images_dir: Path) -> list[Path]:
    seen: list[Path] = []
    names = set()
    for match in _IMG_MD.finditer(md):
        name = Path(match.group(2).replace("\\", "/")).name
        path = images_dir / name
        if name not in names and path.is_file():
            names.add(name)
            seen.append(path)
    for extra in sorted(images_dir.glob("*")):
        if extra.suffix.lower() in {".png", ".jpg", ".jpeg", ".gif", ".webp"} and extra not in seen:
            seen.append(extra)
    return seen


def _pdf_image_count(pdf_path: Path) -> int:
    doc = pymupdf.open(pdf_path)
    try:
        return sum(len(page.get_images()) for page in doc)
    finally:
        doc.close()


def _text_without_images(md: str) -> str:
    return re.sub(r"!\[.*?\]\(.*?\)", "", md).strip()


def _pdf_has_text_layer(pdf_path: Path) -> bool:
    doc = pymupdf.open(pdf_path)
    try:
        text = "".join((page.get_text() or "") for page in doc)
        return len(re.sub(r"\s+", "", text)) >= 10
    finally:
        doc.close()


def _try_marker(pdf_path: Path, out_dir: Path) -> str | None:
    marker = shutil.which("marker_single") or shutil.which("marker")
    if not marker:
        return None
    proc = subprocess.run(
        [marker, str(pdf_path), "--output_dir", str(out_dir)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.returncode != 0:
        return None
    mds = list(out_dir.rglob("*.md"))
    if mds:
        return mds[0].read_text(encoding="utf-8")
    return None


def convert_pdf(pdf_path: Path, out_dir: Path) -> ConvertResult:
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    images = out_dir / "images"
    images.mkdir(parents=True, exist_ok=True)

    md = pymupdf4llm.to_markdown(
        str(pdf_path),
        write_images=True,
        image_path=str(images),
    )
    md = _rewrite_image_links(md, images)
    stripped = _text_without_images(md)
    empty_text = len(re.sub(r"\s+", "", stripped)) < 40 and not _pdf_has_text_layer(pdf_path)
    if empty_text:
        marker_md = _try_marker(pdf_path, out_dir)
        if marker_md is None:
            raise ConvertError("Нужен текстовый PDF или Marker")
        md = _rewrite_image_links(marker_md, images)
        stripped = _text_without_images(md)
        empty_text = len(stripped) < 40

    image_paths = _image_paths_in_md_order(md, images)
    pdf_n = _pdf_image_count(pdf_path)
    warning = None
    if pdf_n > len(image_paths):
        warning = f"В PDF больше изображений, чем извлечено: {pdf_n} vs {len(image_paths)}"
        print(warning, file=sys.stderr)

    md_path = out_dir / "assignment.md"
    md_path.write_text(md, encoding="utf-8")
    return ConvertResult(
        md_path=md_path,
        image_paths=image_paths,
        warning=warning,
        empty_text=empty_text,
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    try:
        result = convert_pdf(Path(args.pdf), Path(args.out))
    except ConvertError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(result.md_path)
    if result.warning:
        print(result.warning, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
