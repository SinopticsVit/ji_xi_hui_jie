#!/usr/bin/env python3
"""Generate PDF from Russian markdown translation."""
from pathlib import Path
import re
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BASE = Path(__file__).parent
MD_PATH = BASE / "网络竞价方案-ru.md"
PDF_PATH = BASE / "网络竞价方案-ru.pdf"
FONT_REG = "Arial"
FONT_BOLD = "Arial-Bold"

pdfmetrics.registerFont(TTFont(FONT_REG, r"C:\Windows\Fonts\arial.ttf"))
pdfmetrics.registerFont(TTFont(FONT_BOLD, r"C:\Windows\Fonts\arialbd.ttf"))


def md_inline_to_xml(text: str) -> str:
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", rf"<b>\1</b>", text)
    text = re.sub(r"`(.+?)`", rf"<font name='Courier'>\1</font>", text)
    text = re.sub(r"\[(.+?)\]\((.+?)\)", r"<link href='\2' color='blue'>\1</link>", text)
    return text


def build_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "title",
            parent=styles["Normal"],
            fontName=FONT_BOLD,
            fontSize=16,
            leading=20,
            alignment=TA_CENTER,
            spaceAfter=10,
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            parent=styles["Normal"],
            fontName=FONT_REG,
            fontSize=12,
            leading=16,
            alignment=TA_CENTER,
            spaceAfter=6,
        ),
        "h2": ParagraphStyle(
            "h2",
            parent=styles["Normal"],
            fontName=FONT_BOLD,
            fontSize=13,
            leading=17,
            spaceBefore=12,
            spaceAfter=8,
        ),
        "h3": ParagraphStyle(
            "h3",
            parent=styles["Normal"],
            fontName=FONT_BOLD,
            fontSize=11.5,
            leading=15,
            spaceBefore=10,
            spaceAfter=6,
        ),
        "h4": ParagraphStyle(
            "h4",
            parent=styles["Normal"],
            fontName=FONT_BOLD,
            fontSize=11,
            leading=14,
            spaceBefore=8,
            spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "body",
            parent=styles["Normal"],
            fontName=FONT_REG,
            fontSize=10.5,
            leading=14,
            alignment=TA_JUSTIFY,
            spaceAfter=6,
        ),
        "bullet": ParagraphStyle(
            "bullet",
            parent=styles["Normal"],
            fontName=FONT_REG,
            fontSize=10.5,
            leading=14,
            leftIndent=14,
            bulletIndent=0,
            spaceAfter=4,
        ),
        "meta": ParagraphStyle(
            "meta",
            parent=styles["Normal"],
            fontName=FONT_REG,
            fontSize=9,
            leading=12,
            textColor=colors.grey,
            alignment=TA_CENTER,
            spaceBefore=16,
        ),
    }


def parse_table(lines: list[str], start: int) -> tuple[list[list[str]], int]:
    rows = []
    i = start
    while i < len(lines) and lines[i].strip().startswith("|"):
        row = [cell.strip() for cell in lines[i].strip().strip("|").split("|")]
        if not all(set(cell) <= {"-", ":"} for cell in row):
            rows.append(row)
        i += 1
    return rows, i


def build_story(text: str, styles: dict):
    story = []
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        if stripped == "---":
            story.append(Spacer(1, 0.15 * cm))
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
            story.append(Spacer(1, 0.15 * cm))
            i += 1
            continue

        if stripped.startswith("# "):
            story.append(Paragraph(md_inline_to_xml(stripped[2:]), styles["title"]))
            i += 1
            continue

        if stripped.startswith("## "):
            story.append(Paragraph(md_inline_to_xml(stripped[3:]), styles["h2"]))
            i += 1
            continue

        if stripped.startswith("### "):
            story.append(Paragraph(md_inline_to_xml(stripped[4:]), styles["h3"]))
            i += 1
            continue

        if stripped.startswith("#### "):
            story.append(Paragraph(md_inline_to_xml(stripped[5:]), styles["h4"]))
            i += 1
            continue

        if stripped.startswith("|"):
            rows, i = parse_table(lines, i)
            if rows:
                table = Table(
                    [[Paragraph(md_inline_to_xml(c), styles["body"]) for c in row] for row in rows],
                    colWidths=[5.5 * cm, 11 * cm],
                    repeatRows=1,
                )
                table.setStyle(
                    TableStyle(
                        [
                            ("FONTNAME", (0, 0), (-1, -1), FONT_REG),
                            ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0f0f0")),
                            ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
                            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ]
                    )
                )
                story.append(Spacer(1, 0.1 * cm))
                story.append(table)
                story.append(Spacer(1, 0.2 * cm))
            continue

        if stripped.startswith("- "):
            story.append(
                Paragraph(
                    f"• {md_inline_to_xml(stripped[2:])}",
                    styles["bullet"],
                )
            )
            i += 1
            continue

        if re.match(r"^\d+\.\s", stripped):
            story.append(Paragraph(md_inline_to_xml(stripped), styles["body"]))
            i += 1
            continue

        if stripped.startswith("*") and stripped.endswith("*"):
            story.append(Paragraph(md_inline_to_xml(stripped.strip("*")), styles["meta"]))
            i += 1
            continue

        story.append(Paragraph(md_inline_to_xml(stripped), styles["body"]))
        i += 1

    return story


def main():
    text = MD_PATH.read_text(encoding="utf-8")
    styles = build_styles()
    doc = SimpleDocTemplate(
        str(PDF_PATH),
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title="Программа сетевых торгов B2409",
        author="Translation",
    )
    doc.build(build_story(text, styles))
    print(f"Created: {PDF_PATH}")


if __name__ == "__main__":
    main()
