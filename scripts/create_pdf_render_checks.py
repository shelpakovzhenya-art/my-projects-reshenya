from pathlib import Path
import re

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer
import pypdfium2 as pdfium

from create_seo_reports import REPORTS, ROOT, why_for


PDF_DIR = ROOT / "reports-pdf"
PNG_DIR = ROOT / "render-check-png"
FONT_DIR = Path("C:/Windows/Fonts")


def register_fonts():
    regular = FONT_DIR / "arial.ttf"
    bold = FONT_DIR / "arialbd.ttf"
    if regular.exists():
        pdfmetrics.registerFont(TTFont("ReportRegular", str(regular)))
    else:
        pdfmetrics.registerFont(TTFont("ReportRegular", str(FONT_DIR / "segoeui.ttf")))
    if bold.exists():
        pdfmetrics.registerFont(TTFont("ReportBold", str(bold)))
    else:
        pdfmetrics.registerFont(TTFont("ReportBold", str(regular)))


def esc(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "ReportTitle",
            parent=base["Title"],
            fontName="ReportBold",
            fontSize=20,
            leading=24,
            textColor=colors.HexColor("#1F4D78"),
            spaceAfter=18,
            alignment=0,
        ),
        "h1": ParagraphStyle(
            "ReportH1",
            parent=base["Heading1"],
            fontName="ReportBold",
            fontSize=16,
            leading=20,
            textColor=colors.HexColor("#2E74B5"),
            spaceBefore=14,
            spaceAfter=8,
        ),
        "body": ParagraphStyle(
            "ReportBody",
            parent=base["BodyText"],
            fontName="ReportRegular",
            fontSize=10.5,
            leading=14,
            spaceAfter=5,
        ),
        "bullet": ParagraphStyle(
            "ReportBullet",
            parent=base["BodyText"],
            fontName="ReportRegular",
            fontSize=10.5,
            leading=14,
            leftIndent=18,
            firstLineIndent=-10,
            spaceAfter=4,
        ),
        "taskHead": ParagraphStyle(
            "TaskHead",
            parent=base["BodyText"],
            fontName="ReportBold",
            fontSize=11,
            leading=15,
            spaceBefore=8,
            spaceAfter=2,
        ),
        "taskBody": ParagraphStyle(
            "TaskBody",
            parent=base["BodyText"],
            fontName="ReportRegular",
            fontSize=10.5,
            leading=14,
            leftIndent=18,
            spaceAfter=2,
        ),
    }


def build_pdf(report, style_map):
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    path = PDF_DIR / f"SEO ТЗ - {report['project']}.pdf"
    doc = SimpleDocTemplate(
        str(path),
        pagesize=letter,
        rightMargin=inch,
        leftMargin=inch,
        topMargin=inch,
        bottomMargin=inch,
        title=f"SEO ТЗ - {report['project']}",
    )
    story = [Paragraph(f"SEO ТЗ: {esc(report['project'])}", style_map["title"])]
    story.append(Paragraph("Сделано", style_map["h1"]))
    for item in report["done"]:
        story.append(Paragraph(f"• {esc(item)}", style_map["bullet"]))

    story.append(Spacer(1, 6))
    story.append(Paragraph("Что нужно сделать", style_map["h1"]))
    for index, (_priority, title, action) in enumerate(report["tz"], start=1):
        story.append(KeepTogether([
            Paragraph(f"{index}. {esc(title)}", style_map["taskHead"]),
            Paragraph(f"<b>Что сделать:</b> {esc(action)}", style_map["taskBody"]),
            Paragraph(f"<b>Зачем:</b> {esc(why_for(title, action))}", style_map["taskBody"]),
        ]))

    doc.build(story)
    return path


def render_pdf_to_png(pdf_path):
    PNG_DIR.mkdir(parents=True, exist_ok=True)
    stem = pdf_path.stem
    for old in PNG_DIR.glob(f"{stem}-page-*.png"):
        old.unlink()
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for index in range(len(pdf)):
            page = pdf[index]
            bitmap = page.render(scale=2).to_pil()
            out = PNG_DIR / f"{stem}-page-{index + 1}.png"
            bitmap.save(out)
    finally:
        pdf.close()


def main():
    register_fonts()
    style_map = styles()
    for report in REPORTS:
        pdf_path = build_pdf(report, style_map)
        render_pdf_to_png(pdf_path)
        print(pdf_path)


if __name__ == "__main__":
    main()
