#!/usr/bin/env python3
"""Split catalog grid slides into per-SKU PNG thumbnails."""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
MATERIALS_DIR = HERE / "maretials"
YUHUAN_DIR = MATERIALS_DIR / "20260406-yuhuan-tianrun"
IMAGES_DIR = YUHUAN_DIR / "images"
CATALOG_MD = MATERIALS_DIR / "catalog-ru.md"
CATALOG_ITEMS_DIR = IMAGES_DIR / "catalog-items"

# Calibrated for 1500x844 slide PNGs (150 DPI render).
GRID_LEFT = 32
GRID_RIGHT = 1478
GRID_TOP = 106
GRID_BOTTOM = 828

SLIDE10_ROW_COLS = [9, 9, 9, 9, 9, 7]
SLIDE11_ROW_COLS = [9, 9, 12]

IMAGE_PREFIX = "20260406-yuhuan-tianrun/images/catalog-items"


@dataclass(frozen=True)
class CatalogItem:
    slide: int
    sku: str
    index: int


def sanitize_filename(sku: str) -> str:
    return re.sub(r"\s+", "_", sku.strip())


def sku_to_relpath(slide: int, sku: str) -> str:
    name = sanitize_filename(sku)
    return f"{IMAGE_PREFIX}/slide-{slide}/{name}.png"


def parse_catalog_items(text: str) -> tuple[list[str], list[str]]:
    start = text.index("#### A.2.1")
    mid = text.index("#### A.2.2")
    end = text.index("#### A.2.3")
    block10 = text[start:mid]
    block11 = text[mid:end]

    def extract_skus(block: str) -> list[str]:
        return [m.group(1).strip() for m in re.finditer(r"^\| \d+ \| ([^|]+) \|", block, re.M)]

    return extract_skus(block10), extract_skus(block11)


def cell_boxes(slide: int) -> list[tuple[int, int, int, int]]:
    row_cols = SLIDE10_ROW_COLS if slide == 10 else SLIDE11_ROW_COLS
    grid_w = GRID_RIGHT - GRID_LEFT
    grid_h = GRID_BOTTOM - GRID_TOP
    row_h = grid_h / len(row_cols)
    boxes: list[tuple[int, int, int, int]] = []

    for row_idx, cols in enumerate(row_cols):
        col_w = grid_w / cols
        y0 = int(round(GRID_TOP + row_idx * row_h))
        y1 = int(round(GRID_TOP + (row_idx + 1) * row_h))
        for col_idx in range(cols):
            x0 = int(round(GRID_LEFT + col_idx * col_w))
            x1 = int(round(GRID_LEFT + (col_idx + 1) * col_w))
            boxes.append((x0, y0, x1, y1))
    return boxes


def split_slide(slide: int, skus: list[str], force: bool) -> list[dict]:
    src = IMAGES_DIR / f"slide-{slide:02d}.png"
    if not src.is_file():
        raise FileNotFoundError(src)

    boxes = cell_boxes(slide)
    if len(skus) != len(boxes):
        raise ValueError(f"slide-{slide}: {len(skus)} SKUs vs {len(boxes)} grid cells")

    out_dir = CATALOG_ITEMS_DIR / f"slide-{slide}"
    out_dir.mkdir(parents=True, exist_ok=True)

    img = Image.open(src)
    manifest_rows: list[dict] = []

    for idx, (sku, box) in enumerate(zip(skus, boxes)):
        row_idx = 0
        col_idx = idx
        row_cols = SLIDE10_ROW_COLS if slide == 10 else SLIDE11_ROW_COLS
        consumed = 0
        for r, cols in enumerate(row_cols):
            if idx < consumed + cols:
                row_idx = r
                col_idx = idx - consumed
                break
            consumed += cols

        filename = f"{sanitize_filename(sku)}.png"
        out_path = out_dir / filename
        if out_path.exists() and not force:
            pass
        else:
            crop = img.crop(box)
            crop.save(out_path, format="PNG", optimize=True)

        rel = f"catalog-items/slide-{slide}/{filename}"
        manifest_rows.append(
            {
                "slide": slide,
                "index": idx + 1,
                "row": row_idx + 1,
                "col": col_idx + 1,
                "sku": sku,
                "filename": filename,
                "path": rel,
                "bbox": list(box),
            }
        )
    return manifest_rows


def update_catalog_md(force: bool) -> None:
    text = CATALOG_MD.read_text(encoding="utf-8")
    start = text.index("#### A.2.1")
    end = text.index("#### A.2.3")
    head = text[:start]
    block = text[start:end]
    tail = text[end:]

    skus10, skus11 = parse_catalog_items(text)
    sku_slide: dict[str, int] = {s: 10 for s in skus10}
    sku_slide.update({s: 11 for s in skus11})

    def thumb_cell(sku: str) -> str:
        slide = sku_slide[sku.strip()]
        rel = sku_to_relpath(slide, sku.strip())
        alt = sanitize_filename(sku)
        return f"![{alt}]({rel})"

    def clean_note(note: str) -> str:
        note = note.strip()
        replacements = {
            "см. миниатюру": "—",
            "см. миниатюру на слайде": "—",
        }
        for old, new in replacements.items():
            if note == old:
                return "—"
            note = note.replace(f", {old}", "").replace(old, "—")
        note = note.replace(", —", "").strip()
        return note or "—"

    lines = block.splitlines()
    out_lines: list[str] = []
    for line in lines:
        if line.startswith("| № | Артикул |") and "Миниатюра" not in line:
            out_lines.append("| № | Артикул | Миниатюра | Наименование / тип | Примечание |")
            continue
        if re.match(r"^\|---\|", line) and out_lines and "Миниатюра" in out_lines[-1]:
            out_lines.append("|---|---------|-----------|-------------------|------------|")
            continue
        m = re.match(r"^\| (\d+) \| ([^|]+) \| ([^|]+) \| (.+) \|$", line)
        if m and "Миниатюра" not in line:
            num, sku, name, note = m.group(1), m.group(2).strip(), m.group(3).strip(), m.group(4).strip()
            if sku != "Артикул":
                out_lines.append(
                    f"| {num} | {sku} | {thumb_cell(sku)} | {name} | {clean_note(note)} |"
                )
                continue
        out_lines.append(line)

    new_block = "\n".join(out_lines)
    if new_block == block and not force:
        print("[info] catalog-ru.md already has thumbnail column")
    else:
        CATALOG_MD.write_text(head + new_block + tail, encoding="utf-8")
        print(f"[done] updated {CATALOG_MD}")


def update_index(text: str) -> str:
    text = text.replace(
        "| A.2.1 | `slide-10.png` | Запирающие системы — **52 позиции** (таблицы HB / TR-K) |",
        "| A.2.1 | `slide-10.png` + **52** item PNG | Запирающие системы (таблицы HB / TR-K) |",
    )
    text = text.replace(
        "| A.2.2 | `slide-11.png` | GSE и комплекты — **30 позиций** (таблицы TR-C/D/S/B/H/Q) |",
        "| A.2.2 | `slide-11.png` + **30** item PNG | GSE и комплекты (таблицы TR-C/D/S/B/H/Q) |",
    )
    text = text.replace(
        "**Всего иллюстраций в каталоге: 32** · **позиций в таблицах A.2.1–A.2.2: 82**",
        "**Всего иллюстраций в каталоге: 32** · **позиций A.2.1–A.2.2: 82** · **item PNG: 82** (`catalog-items/`)",
    )
    return text


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--slide", choices=["10", "11", "all"], default="all")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--skip-catalog", action="store_true")
    args = ap.parse_args()

    catalog_text = CATALOG_MD.read_text(encoding="utf-8")
    skus10, skus11 = parse_catalog_items(catalog_text)

    manifest: list[dict] = []
    slides = [10, 11] if args.slide == "all" else [int(args.slide)]

    for slide in slides:
        skus = skus10 if slide == 10 else skus11
        rows = split_slide(slide, skus, force=args.force)
        manifest.extend(rows)
        print(f"[done] slide-{slide}: {len(rows)} thumbnails -> {CATALOG_ITEMS_DIR / f'slide-{slide}'}")

    manifest_path = CATALOG_ITEMS_DIR / "manifest.json"
    if args.slide == "all":
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        existing: list[dict] = []
        if manifest_path.is_file():
            existing = json.loads(manifest_path.read_text(encoding="utf-8"))
        existing = [r for r in existing if r["slide"] != int(args.slide)]
        existing.extend(manifest)
        existing.sort(key=lambda r: (r["slide"], r["index"]))
        manifest_path.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[done] manifest: {manifest_path} ({len(manifest) if args.slide != 'all' else len(manifest)} entries this run)")

    if not args.skip_catalog:
        update_catalog_md(force=args.force)
        text = CATALOG_MD.read_text(encoding="utf-8")
        CATALOG_MD.write_text(update_index(text), encoding="utf-8")

    return 0


if __name__ == "__main__":
    sys.exit(main())
