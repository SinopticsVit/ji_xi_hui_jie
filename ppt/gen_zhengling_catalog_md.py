#!/usr/bin/env python3
"""Generate zhengling-machinery/catalog-ru.md from content-ru.md."""

from __future__ import annotations

import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
Z_DIR = HERE / "maretials" / "zhengling-machinery"
CONTENT = Z_DIR / "content-ru.md"
OUTPUT = Z_DIR / "catalog-ru.md"

SECTIONS = [
    ("C.1", "Профиль и компетенции", [
        ("C.1.1", "Профиль компании", 1),
        ("C.1.2", "Обзор предприятия", 2),
        ("C.1.3", "Команда НИОКР", 3),
        ("C.1.4", "Динамика продаж и прогноз", 4),
        ("C.1.5", "Корпоративный дух", 5),
        ("C.1.6", "Принципы управления", 6),
        ("C.1.7", "Сертификаты и лицензии", 7),
        ("C.1.8", "Награды предприятия", 8),
        ("C.1.9", "Обзор (слайд 9)", 9),
    ]),
    ("C.2", "Производственная база", [
        ("C.2.1", "Механообрабатывающий цех", 10),
        ("C.2.2", "Линии сборки I-Shaft", 11),
        ("C.2.3", "Контрольно-измерительное оборудование", 12),
    ]),
    ("C.3", "Продукция", [
        ("C.3.1", "Yoke (вилки)", 13),
        ("C.3.2", "Линейка I-Shaft", 14),
        ("C.3.3", "10T скользящий промежуточный вал", 15),
        ("C.3.4", "18T скользящий промежуточный вал", 16),
        ("C.3.5", "EPS ball-sliding I-Shaft", 17),
        ("C.3.6", "Рулевые валы", 18),
    ]),
    ("C.4", "Проектирование и валидация", [
        ("C.4.1", "Карданный шарнир и фазировка", 19),
        ("C.4.2", "FEA-анализ", 20),
        ("C.4.3", "Ключевые характеристики", 21),
        ("C.4.4", "Требования к материалам", 22),
        ("C.4.5", "Программа испытаний I-Shaft", 23),
    ]),
    ("C.5", "Производство и контроль", [
        ("C.5.1", "Сварка yoke I-Shaft", 24),
        ("C.5.2", "Производство (слайд 25)", 25),
        ("C.5.3", "Spline OM I-Shaft", 26),
        ("C.5.4", "Ball Sliding I-Shaft", 27),
        ("C.5.5", "Сборка I-Shaft", 28),
        ("C.5.6", "Мониторинг установки подшипников", 29),
        ("C.5.7", "Контроль линии карданных шарниров", 30),
        ("C.5.8", "Стенды контроля карданных шарниров", 31),
        ("C.5.9", "Испытание жёсткости и зазоров", 32),
    ]),
    ("C.6", "Клиенты и сервис", [
        ("C.6.1", "Клиенты по регионам", 33),
        ("C.6.2", "Дух обслуживания", 34),
        ("C.6.3", "Контакты", 35),
    ]),
]

SKIP_PATTERNS = [
    r"^Yongjia County Zhengling.*$",
    r"^永嘉县正菱.*$",
    r"^Zhengling Machinery.*$",
    r"^Company Profile$",
    r"^R & D team$",
    r"^SALES PERFORMANCE.*$",
    r"^The company spirit.*$",
    r"^Like Steering System.*$",
    r"^Technology innovation.*$",
    r"^Management Concepts$",
    r"^Work by wistom.*$",
    r"^MACHINING WORKSHOP$",
    r"^Machining workshop.*$",
    r"^Plant Overview$",
    r"^production Equipments$",
    r"^Intermediate Shaft Assembly$",
    r"^Inspection Equipments$",
    r"^Provide the OEM.*$",
    r"^Product line$",
    r"^Provide the R&D.*$",
    r"^10T slide spline$",
    r"^\(welding yoke.*$",
    r"^Assistance load.*$",
    r"^18T slide spline$",
    r"^EPS sliding connection.*$",
    r"^it has two different.*$",
    r"^This design use.*$",
    r"^This design adopts.*$",
    r"^Union$",
    r"^Steering shafts$",
    r"^I-Shaft yoke welding$",
    r"^Spline OM I-Shaft$",
    r"^Ball Sliding I-Shaft$",
    r"^I-Shaft ASM$",
    r"^Designated Characteristics$",
    r"^Material Requirement.*$",
    r"^Contact us$",
    r"^См\. изображение$",
    r"^\d+$",
    r"^Профиль компании$",
]


def parse_slides(text: str) -> dict[int, str]:
    parts = re.split(r"\n## Слайд (\d+)\n", text)
    slides: dict[int, str] = {}
    for i in range(1, len(parts), 2):
        num = int(parts[i])
        body = parts[i + 1].strip()
        slides[num] = body
    return slides


def extract_body(slide_text: str) -> tuple[str, list[str]]:
    lines = slide_text.splitlines()
    bullets: list[str] = []
    paras: list[str] = []
    in_table = False
    table_lines: list[str] = []

    for line in lines:
        s = line.strip()
        if s.startswith("![") or not s:
            continue
        if s.startswith("|"):
            in_table = True
            table_lines.append(line)
            continue
        if in_table and not s.startswith("|"):
            in_table = False
        if any(re.match(p, s, re.I) for p in SKIP_PATTERNS):
            continue
        if s.startswith("- "):
            bullets.append(s[2:].strip())
        elif s.startswith("  - "):
            bullets.append(s[4:].strip())
        else:
            paras.append(s)

    note = "\n".join(paras)
    if table_lines:
        note = (note + "\n\n" + "\n".join(table_lines)).strip()
    return note, bullets


def main() -> None:
    text = CONTENT.read_text(encoding="utf-8")
    slides = parse_slides(text)

    out: list[str] = [
        "# Каталог продукции и иллюстраций (Zhengling)",
        "",
        "Полный каталог по материалам производственной базы **автомобильных компонентов "
        "рулевой системы**.",
        "",
        "Источник (внутр.): [content-ru.md](content-ru.md)",
        "",
        "Сводка возможностей: [capabilities-zhengling-ru.md](capabilities-zhengling-ru.md)",
        "",
        "---",
        "",
        "## Раздел C. Компоненты рулевой системы (автомобильная отрасль)",
        "",
    ]

    index_rows: list[str] = []

    for block_id, block_title, items in SECTIONS:
        out.append(f"### {block_id} {block_title}")
        out.append("")
        for sub_id, title, slide_num in items:
            body = slides.get(slide_num, "")
            note, bullets = extract_body(body)
            out.append(f"#### {sub_id} {title}")
            out.append("")
            out.append(f"![{title}](images/slide-{slide_num:02d}.png)")
            out.append("")
            if note:
                out.append(note)
                out.append("")
            for b in bullets:
                out.append(f"- {b}")
            if bullets:
                out.append("")
            index_rows.append(
                f"| {sub_id} | `images/slide-{slide_num:02d}.png` | {title} |"
            )

    out.extend([
        "---",
        "",
        "## Индекс иллюстраций",
        "",
        "| ID | Файл | Раздел |",
        "|---|---|---|",
        *index_rows,
        "",
        f"**Всего иллюстраций: {len(index_rows)}**",
        "",
    ])

    OUTPUT.write_text("\n".join(out), encoding="utf-8")
    print(f"[done] {OUTPUT} ({len(index_rows)} sections)")


if __name__ == "__main__":
    main()
