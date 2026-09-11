from pathlib import Path
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def make_two_image_pdf(path: Path) -> None:
    path = Path(path)
    img_dir = path.parent / "_pdf_imgs"
    img_dir.mkdir(parents=True, exist_ok=True)
    red = img_dir / "red.png"
    blue = img_dir / "blue.png"
    Image.new("RGB", (40, 40), (255, 0, 0)).save(red)
    Image.new("RGB", (40, 40), (0, 0, 255)).save(blue)
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFont("Times-Roman", 14)
    c.drawString(72, h - 72, "BEFORE")
    c.drawImage(str(red), 72, h - 160, 40, 40)
    c.drawString(72, h - 200, "MIDDLE")
    c.drawImage(str(blue), 72, h - 280, 40, 40)
    c.drawString(72, h - 320, "AFTER")
    c.save()
