<div align="center">

# 🏗️ FinModel Agent

**ИИ-агент для финансового моделирования складской недвижимости**

[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![Claude Code](https://img.shields.io/badge/Claude_Code-MCP-D4A017?style=flat-square&logo=anthropic&logoColor=white)](https://claude.ai/code)
[![Excel](https://img.shields.io/badge/Excel-openpyxl-217346?style=flat-square&logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)

Агент принимает параметры проекта на языке бизнеса и генерирует полноценную  
Excel-финмодель с формулами, денежными потоками, KPI и анализом чувствительности.

</div>

---

## ✨ Что умеет агент

| Команда | Описание |
|---|---|
| `/new-project` | Диалог: собрать параметры → сохранить `projects/<name>.json` |
| `/generate-report <name>` | Построить Excel-модель одним MCP-вызовом (4 листа, ~2700 формул) |
| `/analyze <name>` | Прочитать KPI → аналитическое заключение на языке бизнеса |
| `/sensitivity <name>` | Обновить лист «Чувствительность» с новыми сценариями |
| `/load-excel <путь>` | Загрузить параметры из существующего Excel-шаблона |

### Что внутри сгенерированного Excel

```
📊 reports/<name>_finmodel.xlsx
├── Параметры        — входные данные + производные показатели (синие ячейки)
├── Денежные потоки  — помесячный CF на весь горизонт (до 180 строк формул)
├── ИТОГ             — NPV, IRR, ROI, срок окупаемости
└── Чувствительность — сценарии ±20% по 5 ключевым параметрам
```

**Всё в формулах.** Меняйте любую синюю ячейку на листе «Параметры» — модель пересчитывается автоматически.

---

## 🚀 Быстрый старт

### 1. Установка зависимостей

```bash
pip install -r requirements.txt
```

### 2. Настройка MCP-серверов

Обновите `.mcp.json` — укажите путь к проекту на вашем компьютере:

```json
{
  "mcpServers": {
    "finmodel": {
      "type": "stdio",
      "command": "python",
      "args": ["-m", "finmodel.mcp_server"],
      "cwd": "/path/to/diplom2"
    },
    "excel": {
      "type": "stdio",
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio"],
      "env": {
        "EXCEL_FILES_PATH": "/path/to/diplom2"
      }
    }
  }
}
```

### 3. Откройте проект в Claude Code

```bash
cd diplom2
claude
```

### 4. Запустите первый проект

```
/new-project
```

Агент задаст вопросы и создаст параметры. Затем:

```
/generate-report <name>
```

Excel-модель появится в папке `reports/`.

---

## 🧮 Финансовая логика

Агент моделирует **две фазы**:

**Фаза строительства** (месяцы 1..Tc)
- Равномерное освоение капвложений и кредитных траншей
- Начисление процентов на накопленный остаток кредита
- Отрицательный CF инвестора

**Фаза эксплуатации** (месяцы Tc+1..Th)
- Выручка с ежегодной индексацией на инфляцию
- Аннуитетное погашение кредита
- НДС: возмещение строительного НДС в первые 3 месяца
- Налог на прибыль, налог на имущество, амортизация
- Положительный CF инвестора

**KPI на листе ИТОГ:**

| Показатель | Формула |
|---|---|
| NPV | Дисконтированная сумма CF инвестора по ставке ЦБ/12 |
| IRR | Годовая внутренняя норма доходности |
| ROI | Сумма чистой прибыли / Собственные средства |
| Payback | Первый месяц с накопленным CF ≥ 0 |

---

## 🗂️ Структура проекта

```
diplom2/
├── finmodel/               # MCP-сервер и расчётный движок
│   ├── mcp_server.py       # FastMCP: list/get/build/compare/cashflow
│   ├── build_excel.py      # Генератор Excel (4 листа, формулы)
│   ├── finmodel.py         # Расчёт CF, KPI, амортизации, налогов
│   └── config.py           # Пути к projects/ и reports/
├── .claude/
│   └── skills/             # Скиллы Claude Code (/new-project, /generate-report…)
├── projects/               # JSON-параметры проектов
├── reports/                # Готовые Excel-финмодели
├── templates/              # Базовые шаблоны
├── memory.md               # Память агента (пользователь, компания, дефолты)
├── CLAUDE.md               # Инструкции агента
├── .mcp.json               # Конфигурация MCP-серверов
└── requirements.txt
```

---

## ⚙️ MCP-инструменты сервера `finmodel`

| Инструмент | Что делает |
|---|---|
| `list_projects` | Список проектов в `projects/` |
| `get_project_params(name)` | Параметры проекта (сырой JSON) |
| `get_kpi(name)` | NPV / IRR / ROI / Payback без Excel |
| `build_project(name)` | **Полный отчёт** + KPI одним вызовом |
| `compare_projects([a, b])` | Сравнение KPI нескольких проектов |
| `get_cashflow_summary(name)` | Первые и последние N строк CF |

---

## 🏭 Рыночные ориентиры (Россия 2026)

| Параметр | Мин. | Типовой | Макс. |
|---|---|---|---|
| Стоимость строительства, руб./м² с НДС | 40 000 | 60 000 | 90 000 |
| Аренда склада A, руб./м²/мес. без НДС | 800 | 1 100 | 1 500 |
| Индексация аренды | 4% | 5–7% | 10% |
| Ставка ЦБ | 12% | 15% | 21% |
| Кредитная ставка (надбавка к ЦБ) | +4 п.п. | +5 п.п. | +8 п.п. |
| Срок строительства | 4 мес. | 6 мес. | 12 мес. |
| Горизонт планирования | 60 мес. | 120 мес. | 180 мес. |

---

## 📦 Зависимости

```
openpyxl>=3.1.0       # Генерация Excel-файлов
numpy-financial>=1.0.0 # IRR (опционально, есть встроенная реализация)
mcp                    # Model Context Protocol (Claude Code)
```

---

<div align="center">

Сделано с помощью [Claude Code](https://claude.ai/code) · MCP · openpyxl

</div>
