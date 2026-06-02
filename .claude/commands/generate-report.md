# /generate-report

Построить Excel-финмодель проекта **одним MCP-вызовом** нашего внутреннего сервера `finmodel`.

**Никаких записей в ячейки, никаких формул в чате, никаких циклов.**
**1 MCP-вызов на проект. ~50 токенов.**

## Запрещено

- Использовать `mcp__excel__*` для построения модели (низкоуровневый сервер)
- Подставлять числа вместо формул
- Симулировать таблицу в чате

## Шаг 1 — Один вызов

```
mcp__finmodel__build_project
  project_name = "<name>"
```

Сервер:
1. Читает `projects/<name>.json` (внутри себя, агенту не нужно)
2. Нормализует параметры (cost_with_vat → construction_cost, equity_share → loan_share, и т.д.)
3. Генерирует `reports/<name>_finmodel.xlsx` через `build_excel.py` (4 листа, ~2700 формул)
4. Считает KPI в памяти и возвращает их

Возвращает JSON:
```json
{
  "status": "ok",
  "file": "reports/<name>_finmodel.xlsx",
  "horizon_months": 120,
  "construction_months": 12,
  "kpi": {
    "npv": ...,
    "irr_annual": ...,
    "roi": ...,
    "payback_months": ...,
    "total_net_profit": ...,
    "total_income_tax": ...,
    "total_vat_paid": ...,
    "total_vat_refund": ...,
    "total_property_tax": ...,
    "total_interest_paid": ...,
    "total_revenue_net": ...,
    "total_revenue_gross": ...,
    "total_operating_costs": ...
  }
}
```

## Шаг 2 — Показать пользователю

Берём KPI прямо из ответа сервера, форматируем числа:

```
✅ Финмодель готова: reports/<name>_finmodel.xlsx

   Все расчёты — формулы Excel. Чтобы изменить параметры,
   откройте лист «Параметры» и измените синие ячейки C4:C17.
   Модель пересчитается автоматически.

   Листы (все заполнены автоматически при сборке):
     • Параметры        — входные данные + производные формулы (C20:C28)
     • Денежные потоки  — помесячный CF (120 строк формул)
     • ИТОГ             — NPV, IRR, ROI, срок окупаемости
     • Чувствительность — сценарии ±20% по 5 параметрам (формулы-ссылки на «Параметры»)

   KPI:
     NPV:                    <kpi.npv> руб.
     IRR (год.):             <kpi.irr_annual>%
     ROI:                    <kpi.roi>%
     Срок окупаемости:       <kpi.payback_months> мес.
     Чистая прибыль:         <kpi.total_net_profit> руб.

   Налоги за горизонт:
     Налог на прибыль:       <kpi.total_income_tax> руб.
     НДС к уплате:           <kpi.total_vat_paid> руб.
     Возврат НДС:            <kpi.total_vat_refund> руб.
     Налог на имущество:     <kpi.total_property_tax> руб.
     Проценты по кредиту:    <kpi.total_interest_paid> руб.
```

## Связанные инструменты finmodel-сервера

| Инструмент | Применение |
|---|---|
| `mcp__finmodel__list_projects` | Список доступных проектов в `projects/` |
| `mcp__finmodel__get_project_params(name)` | Прочитать сырой JSON проекта |
| `mcp__finmodel__get_kpi(name)` | KPI без построения Excel (быстро) |
| `mcp__finmodel__build_project(name)` | **Основной.** Полный отчёт + KPI |
| `mcp__finmodel__compare_projects([a,b])` | Сравнить KPI нескольких проектов |

## Если файл нужно перестроить (Excel держит блокировку)

Закрыть Excel и повторить вызов. Признак блокировки — файл `~$<name>_finmodel.xlsx`.

## Расширение модели

Если нужен новый параметр, влияющий на расчёты:
1. Я редактирую `scripts/finmodel.py` (модель) и/или `scripts/build_excel.py` (раскладка ячеек)
2. Сервер перезапускается автоматически (Claude Code следит за `.mcp.json` и кодом)
3. В `/new-project.md` добавляется вопрос
4. Готово, навсегда — все будущие отчёты учитывают параметр
