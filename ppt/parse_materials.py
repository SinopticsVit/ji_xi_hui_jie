#!/usr/bin/env python3
"""Parse presentation materials in maretials/ to markdown + PNG images."""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import fitz
import pytesseract
from PIL import Image

HERE = Path(__file__).resolve().parent
MATERIALS_DIR = HERE / "maretials"
LOCAL_TESSDATA = Path.home() / "AppData" / "Local" / "tesseract" / "tessdata"
MIN_TEXT_CHARS = 30
OCR_LANG = "chi_sim+eng"

TESSERACT_CANDIDATES = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
]

LIBREOFFICE_CANDIDATES = [
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
]


@dataclass(frozen=True)
class Material:
    key: str
    source_name: str
    output_dir: str


MATERIALS: tuple[Material, ...] = (
    Material("yuhuan", "20260406-玉环天润简介.pdf", "20260406-yuhuan-tianrun"),
    Material("qifeng", "七丰画册 2022.pdf", "qifeng-2022"),
    Material("zhengling", "正菱机械科技股份有限公司简介.ppt", "zhengling-machinery"),
)


def find_executable(candidates: list[Path], name: str) -> Path:
    for path in candidates:
        if path.is_file():
            return path
    found = shutil.which(name)
    if found:
        return Path(found)
    raise FileNotFoundError(
        f"{name} not found. Install it and ensure it is on PATH."
    )


def configure_tesseract() -> None:
    import os

    tesseract = find_executable(TESSERACT_CANDIDATES, "tesseract")
    pytesseract.pytesseract.tesseract_cmd = str(tesseract)
    if LOCAL_TESSDATA.is_dir() and any(LOCAL_TESSDATA.glob("*.traineddata")):
        os.environ["TESSDATA_PREFIX"] = str(LOCAL_TESSDATA)


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in text.split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines)


def looks_like_table(lines: list[str]) -> bool:
    if len(lines) < 2:
        return False
    tab_lines = sum(1 for line in lines if "\t" in line)
    if tab_lines >= 2:
        return True
    numeric_cells = 0
    for line in lines:
        parts = re.split(r"\s{2,}|\t", line.strip())
        if len(parts) >= 3:
            numeric_cells += 1
    return numeric_cells >= 3


def text_to_markdown(text: str) -> str:
    text = normalize_text(text)
    if not text:
        return ""

    lines = text.split("\n")
    if looks_like_table(lines):
        rows: list[list[str]] = []
        for line in lines:
            if not line.strip():
                continue
            if "\t" in line:
                cells = [cell.strip() for cell in line.split("\t")]
            else:
                cells = [cell.strip() for cell in re.split(r"\s{2,}", line.strip())]
            rows.append(cells)
        if rows:
            width = max(len(row) for row in rows)
            rows = [row + [""] * (width - len(row)) for row in rows]
            header = rows[0]
            body = rows[1:] if len(rows) > 1 else []
            md_lines = [
                "| " + " | ".join(header) + " |",
                "| " + " | ".join(["---"] * width) + " |",
            ]
            for row in body:
                md_lines.append("| " + " | ".join(row) + " |")
            return "\n".join(md_lines)

    return text


def render_page_png(page: fitz.Page, dpi: int) -> Image.Image:
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def has_cjk(text: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def ocr_image(image: Image.Image) -> str:
    return pytesseract.image_to_string(image, lang=OCR_LANG)


def choose_page_text(extracted: str, image: Image.Image, skip_ocr: bool) -> str:
    if skip_ocr:
        return extracted
    if extracted and (len(extracted) >= MIN_TEXT_CHARS or has_cjk(extracted)):
        return extracted
    ocr_text = normalize_text(ocr_image(image))
    if len(ocr_text) > len(extracted):
        return ocr_text
    return extracted or ocr_text


def convert_ppt_to_pdf(ppt_path: Path, out_dir: Path) -> Path:
    soffice = find_executable(LIBREOFFICE_CANDIDATES, "soffice")
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(soffice),
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(ppt_path),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.returncode != 0:
        raise RuntimeError(
            "LibreOffice conversion failed.\n"
            f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )
    pdf_path = out_dir / f"{ppt_path.stem}.pdf"
    if not pdf_path.is_file():
        pdfs = sorted(out_dir.glob("*.pdf"))
        if not pdfs:
            raise FileNotFoundError(f"No PDF produced from {ppt_path.name}")
        pdf_path = pdfs[0]
    return pdf_path


def resolve_input_path(material: Material) -> Path:
    source = MATERIALS_DIR / material.source_name
    if not source.is_file():
        raise FileNotFoundError(f"Source not found: {source}")
    if source.suffix.lower() == ".ppt":
        tmp_dir = MATERIALS_DIR / material.output_dir / "_tmp"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        return convert_ppt_to_pdf(source, tmp_dir)
    return source


def parse_pdf(
    pdf_path: Path,
    output_dir: Path,
    source_name: str,
    dpi: int,
    skip_ocr: bool,
) -> None:
    images_dir = output_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    slides: list[str] = []

    for index, page in enumerate(doc, start=1):
        image_name = f"slide-{index:02d}.png"
        image_path = images_dir / image_name
        image = render_page_png(page, dpi)
        image.save(image_path, format="PNG", optimize=True)

        extracted = normalize_text(page.get_text("text"))
        body = choose_page_text(extracted, image, skip_ocr)

        body_md = text_to_markdown(body)
        block = [
            f"## Слайд {index}",
            "",
            f"![Слайд {index}](images/{image_name})",
            "",
        ]
        if body_md:
            block.append(body_md)
        slides.append("\n".join(block))

    doc.close()

    title = Path(source_name).stem
    body = "\n\n---\n\n".join(slides)
    content = (
        f"# {title}\n\n"
        f"Источник: [`{source_name}`](../{source_name})\n\n"
        f"---\n\n"
        f"{body}\n"
    )
    (output_dir / "content.md").write_text(content, encoding="utf-8")


def parse_material(
    material: Material,
    dpi: int,
    skip_ocr: bool,
) -> None:
    output_dir = MATERIALS_DIR / material.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = resolve_input_path(material)
    print(f"[parse] {material.source_name} -> {output_dir.name}/")
    parse_pdf(pdf_path, output_dir, material.source_name, dpi, skip_ocr)
    print(f"[done]  {output_dir / 'content.md'}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--only",
        choices=[m.key for m in MATERIALS],
        help="Parse a single material by key.",
    )
    parser.add_argument("--skip-ocr", action="store_true", help="Skip OCR fallback.")
    parser.add_argument("--dpi", type=int, default=150, help="PNG render DPI.")
    args = parser.parse_args()

    configure_tesseract()

    selected = MATERIALS
    if args.only:
        selected = tuple(m for m in MATERIALS if m.key == args.only)

    for material in selected:
        parse_material(material, dpi=args.dpi, skip_ocr=args.skip_ocr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
