# Спецификация презентации — компоненты рулевой системы (Zhengling)

Анонимная презентация 16:9 на русском языке по материалам `capabilities-zhengling-ru.md` и `catalog-ru.md`.

Целевой результат: **16 слайдов**, без названий поставщика, сертификатов, клиентов, контактов и видимых китайских иероглифов на изображениях.

## 1. Ограничения

| Исключить | Слайды источника | Замена |
|---|---|---|
| Профиль, бренд | 1–2 | Нейтральный title |
| Продажи, корпоративный дух | 4–6 | Не включать |
| Сертификаты, награды | 7–8 | Не включать |
| Служебный / пустой | 9 | Не включать |
| Цеха с брендингом | 10–12 | Текстовые слайды НИОКР / матрица |
| Клиенты, сервис, контакты | 33–35 | Не включать |

Допустимы нейтральные термины: I-Shaft, EPS, FEA, NVH, spline, yoke.

## 2. Design Tokens

Как в `maretials/presentation-spec-ru.md`: navy `#0B2A4A`, accent `#0078A8`, 16:9, Segoe UI, footer «N / 16».

## 3. Whitelist изображений

Пути относительно `maretials/`:

| Файл | Назначение | Обработка |
|---|---|---|
| `zhengling-machinery/images/slide-13.png` | Yoke | bottom_mask 14% |
| `zhengling-machinery/images/slide-14.png` | Линейка I-Shaft | bottom_mask 14% |
| `zhengling-machinery/images/slide-15.png` | 10T I-Shaft | bottom_mask 14% |
| `zhengling-machinery/images/slide-16.png` | 18T I-Shaft | bottom_mask 14% |
| `zhengling-machinery/images/slide-17.png` | EPS ball-sliding | bottom_mask 14% |
| `zhengling-machinery/images/slide-18.png` | Рулевые валы | bottom_mask 14% |
| `zhengling-machinery/images/slide-19.png` | Карданный шарнир | crop top 10% + bottom_mask 14% |
| `zhengling-machinery/images/slide-20.png` | FEA | crop top 10% + bottom_mask 14% |
| `zhengling-machinery/images/slide-23.png` | Испытания (опц.) | bottom_mask 14% |
| `zhengling-machinery/images/slide-24.png` | Сварка yoke | crop product zone + bottom_mask 14% |
| `zhengling-machinery/images/slide-26.png` | Spline OM | crop + bottom_mask 14% |
| `zhengling-machinery/images/slide-27.png` | Ball sliding | crop + bottom_mask 14% |
| `zhengling-machinery/images/slide-28.png` | Сборка I-Shaft | crop + bottom_mask 14% |
| `zhengling-machinery/images/slide-32.png` | Испытание жёсткости | crop + bottom_mask 14% |

Temp-каталог: `zhengling-machinery/_pptx_assets/`.

## 4. Структура слайдов (16)

| № | Макет | Заголовок | Контент / изображение |
|---|---|---|---|
| 1 | title | Компоненты рулевой системы | Подзаголовок: автомобильная отрасль, обобщённый обзор |
| 2 | matrix | Карта возможностей | Карточки: Продукция / НИОКР / Производство / Контроль |
| 3 | section | A — Продукция | Промежуточные валы, yoke, рулевые валы |
| 4 | split | НИОКР и проектирование | Бullets из slide-03, без фото |
| 5 | hero | Линейка промежуточных валов | slide-14 |
| 6 | hero | Yoke (вилки) | slide-13 |
| 7 | hero | 10T скользящий I-Shaft | slide-15, момент 30–50 N·m |
| 8 | hero | 18T скользящий I-Shaft | slide-16, момент 50–65 N·m |
| 9 | split | EPS ball-sliding I-Shaft | slide-17, bullets 65–99 N·m |
| 10 | hero | Рулевые валы | slide-18 |
| 11 | dual | Проектирование и FEA | slide-19 + slide-20 |
| 12 | matrix | Программа испытаний | Таблица из slide-23 (функциональные / характеристики / NVH) |
| 13 | section | B — Производство и контроль | 3 линии сборки, сварка, стенды |
| 14 | hero | Сборка промежуточного вала | slide-28 |
| 15 | matrix | Ключевые характеристики и материалы | slide-21 + slide-22 (текст) |
| 16 | title | НИОКР, качество и следующий шаг | Итоговые bullets |

## 5. Сборка

```bash
cd aero_secondhend/ppt
python build_zhengling_presentation_pptx.py
```

Выход: `maretials/zhengling-machinery/presentation-ru.pptx`.
