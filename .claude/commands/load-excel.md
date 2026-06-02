# /load-excel

Загрузить параметры из существующего Excel-файла и сохранить JSON проекта.

## Синтаксис

```
/load-excel <путь_к_файлу>
```

Поддерживаемые форматы: шаблон «ФМ Индустриальный парк Весна 2.0» и аналогичные.

## Шаг 1 — Прочитать лист «Денежные потоки»

```
mcp__excel__read_data_from_excel
  filepath: "<путь_к_файлу>"
  sheet_name: "Денежные потоки"
  start_cell: "A1"
  end_cell: "H15"
```

Извлечь параметры по ячейкам шаблона Весна 2.0:

| Ячейка | Параметр |
|---|---|
| C2 | construction_cost |
| D3 | loan_share |
| C7 | area_sqm |
| C9 | rental_rate_sqm |
| C10 | construction_months |
| C11 | horizon_months |
| G11 | cbr_rate |
| C13 | loan_rate |
| C14 | inflation_rate |
| G14 | vat_rate |

## Шаг 2 — Запросить недостающие данные

Спросить у пользователя (если не читается из файла):
- **Название проекта** (latin, без пробелов)
- **Дата начала строительства** (ГГГГ-ММ)

## Шаг 3 — Сохранить проект через CLI

```bash
python scripts/main.py new-project \
  --name <NAME> \
  --start-date <YYYY-MM> \
  --area <area_sqm> \
  --cost <construction_cost> \
  --loan-share <loan_share> \
  --cbr-rate <cbr_rate> \
  --loan-rate <loan_rate> \
  --rental-rate <rental_rate_sqm> \
  --inflation <inflation_rate> \
  --construction-months <construction_months> \
  --horizon-months <horizon_months>
```

## Шаг 4 — Результат

```
✅ Параметры загружены из: <файл>
   Сохранено: projects/<name>.json

   Площадь:    <area_sqm> кв.м.
   Стоимость:  <construction_cost> руб. (с НДС)
   Аренда:     <rental_rate_sqm> руб./кв.м./мес.
   Горизонт:   <horizon_months> мес.

→ /use-template <name>  — рассчитать новую модель
→ /analyze <name>       — анализ (если файл уже содержит расчёт)
```
