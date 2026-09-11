from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

import pymupdf


class RenderUnavailable(Exception):
    pass


_SOFFICE_PATHS = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
        return True
    except OSError:
        return False


def _soffice_bin() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for path in _SOFFICE_PATHS:
        if Path(path).is_file():
            return path
    return None


def soffice_available() -> bool:
    return _soffice_bin() is not None


def _docx_to_pdf_word(docx_path: Path, pdf_path: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path.resolve()))
        doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _docx_to_pdf_soffice(docx_path: Path, out_dir: Path) -> Path:
    import subprocess

    bin_path = _soffice_bin()
    if not bin_path:
        raise RenderUnavailable("Нет Word и нет LibreOffice")
    proc = subprocess.run(
        [bin_path, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.returncode != 0:
        raise RenderUnavailable("LibreOffice не смог сконвертировать документ")
    pdf = out_dir / (docx_path.stem + ".pdf")
    if not pdf.is_file():
        raise RenderUnavailable("LibreOffice не создал PDF")
    return pdf


def _pdf_to_pngs(pdf_path: Path, out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    pages: list[Path] = []
    try:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=150)
            dest = out_dir / f"page-{i:02d}.png"
            pix.save(dest)
            if dest.stat().st_size == 0:
                raise RenderUnavailable("Пустой PNG страницы")
            pages.append(dest)
    finally:
        doc.close()
    if not pages:
        raise RenderUnavailable("Не удалось получить страницы")
    return pages


def render_docx(docx_path: Path, out_dir: Path) -> list[Path]:
    docx_path = Path(docx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        pdf_path = tmp_dir / "preview.pdf"
        if word_available():
            _docx_to_pdf_word(docx_path, pdf_path)
        elif soffice_available():
            pdf_path = _docx_to_pdf_soffice(docx_path, tmp_dir)
        else:
            raise RenderUnavailable("Нет Word и нет LibreOffice")
        return _pdf_to_pngs(pdf_path, out_dir)


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("docx")
    p.add_argument("--out-dir", required=True)
    args = p.parse_args()
    try:
        pages = render_docx(Path(args.docx), Path(args.out_dir))
    except RenderUnavailable as exc:
        print(str(exc), file=sys.stderr)
        return 1
    for page in pages:
        print(page)
    return 0


if __name__ == "__main__":
    sys.exit(main())
