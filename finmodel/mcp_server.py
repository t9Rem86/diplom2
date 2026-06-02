# -*- coding: utf-8 -*-
"""
MCP-сервер финансовой модели склада.

Тонкая обёртка над scripts/finmodel.py + scripts/build_excel.py.
Агент делает ОДИН вызов build_project → получает Excel + KPI.

Запуск (Claude Code запускает сам через .mcp.json):
    python -m scripts.mcp_server
"""
from __future__ import annotations
import sys
import json
from pathlib import Path

# Корень проекта в sys.path
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from mcp.server.fastmcp import FastMCP

from finmodel.config import PROJECTS_DIR, REPORTS_DIR
from finmodel.finmodel import ProjectParams, calculate, calculate_kpi
from finmodel.build_excel import build


# ─── Маппинг полей JSON → ProjectParams ─────────────────────────────────────
#
# В projects/*.json исторически используются «русские» имена (cost_with_vat,
# equity_share, rent_rate_ex_vat, key_rate, rent_indexation, income_tax, ...).
# ProjectParams хочет старые имена (construction_cost, loan_share, cbr_rate, ...).
# Переводим здесь, в одном месте.

def _parse_start_date(s: str) -> str:
    """
    Привести дату к формату «ГГГГ-ММ» для ProjectParams.
    Понимает «ДД.ММ.ГГГГ» (новый формат), «ГГГГ-ММ-ДД» (старый ISO)
    и уже-нормализованное «ГГГГ-ММ».
    """
    s = s.strip()
    if "." in s:
        # "01.07.2026" → "2026-07"
        parts = s.split(".")
        if len(parts) == 3:
            day, month, year = parts
            return f"{int(year):04d}-{int(month):02d}"
    if "-" in s:
        # "2026-07-01" или "2026-07" → "2026-07"
        parts = s.split("-")
        return f"{int(parts[0]):04d}-{int(parts[1]):02d}"
    raise ValueError(f"Не распознан формат даты: {s!r}")


def _normalize_params(d: dict) -> dict:
    """Преобразовать JSON-словарь в kwargs для ProjectParams."""
    vat = d.get("vat_rate", 0.22)
    out = {
        "name":                d["name"],
        "start_date":          _parse_start_date(d["start_date"]),
        "area_sqm":            float(d["area_sqm"]),
        "construction_cost":   float(d["cost_with_vat"]),
        "loan_share":          1.0 - float(d.get("equity_share", 1 - d.get("debt_share", 0.8))),
        "cbr_rate":            float(d["key_rate"]),
        "loan_rate":           float(d["loan_rate"]),
        # Аренда: в JSON хранится без НДС, ProjectParams ждёт с НДС
        "rental_rate_sqm":     float(d["rent_rate_ex_vat"]) * (1.0 + vat),
        "inflation_rate":      float(d["rent_indexation"]),
        "construction_months": int(d["construction_months"]),
        "horizon_months":      int(d["horizon_months"]),
        "vat_rate":            vat,
        "income_tax_rate":     float(d.get("income_tax", 0.20)),
        "property_tax_rate":   float(d.get("property_tax", 0.022)),
        "operating_cost_pct":  float(d.get("opex_pct", 0.15)),
        "useful_life_years":   int(d.get("useful_life_years", 40)),
    }
    return out


def _load_params(project_name: str) -> ProjectParams | None:
    path = PROJECTS_DIR / f"{project_name}.json"
    if not path.exists():
        return None
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    return ProjectParams(**_normalize_params(raw))


# ─── MCP-сервер ──────────────────────────────────────────────────────────────

mcp = FastMCP("finmodel")


@mcp.tool()
def list_projects() -> str:
    """Список доступных проектов в папке projects/."""
    files = sorted(PROJECTS_DIR.glob("*.json"))
    return json.dumps(
        {"projects": [f.stem for f in files]},
        ensure_ascii=False, indent=2
    )


@mcp.tool()
def get_project_params(project_name: str) -> str:
    """Прочитать сырой JSON проекта по имени."""
    path = PROJECTS_DIR / f"{project_name}.json"
    if not path.exists():
        return json.dumps({"error": f"Проект не найден: {project_name}"}, ensure_ascii=False)
    with open(path, encoding="utf-8") as f:
        return f.read()


@mcp.tool()
def get_kpi(project_name: str) -> str:
    """
    Получить KPI проекта (NPV, IRR, ROI, срок окупаемости, налоги, выручка)
    без построения Excel. Считает в памяти.
    """
    params = _load_params(project_name)
    if params is None:
        return json.dumps({"error": f"Проект не найден: {project_name}"}, ensure_ascii=False)
    rows = calculate(params)
    kpi = calculate_kpi(params, rows)
    return json.dumps(kpi, ensure_ascii=False, indent=2)


@mcp.tool()
def build_project(project_name: str) -> str:
    """
    Построить полную Excel-финмодель проекта.

    Один вызов = готовый отчёт: листы Параметры, Денежные потоки, ИТОГ,
    Чувствительность, все формулы пересчитываются при изменении входных
    данных. Чувствительность встроена автоматически.

    Возвращает путь к файлу и KPI.
    """
    params = _load_params(project_name)
    if params is None:
        return json.dumps({"error": f"Проект не найден: {project_name}"}, ensure_ascii=False)

    rows = calculate(params)
    kpi = calculate_kpi(params, rows)
    out_path = build(params)

    return json.dumps({
        "status": "ok",
        "file": str(out_path),
        "horizon_months": params.horizon_months,
        "construction_months": params.construction_months,
        "kpi": kpi,
        "note": "Все ячейки — формулы Excel. Откройте файл — Excel пересчитает."
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def get_cashflow_summary(project_name: str, n_first: int = 12, n_last: int = 12) -> str:
    """
    Получить первые n_first и последние n_last строк помесячного CF проекта.

    Удобно для /analyze: видны фаза строительства, выход на операционку
    в начале и финальный накопленный CF в конце. Не требует чтения Excel —
    считает в памяти из projects/<name>.json.

    Каждая строка содержит ключевые показатели: дата, фаза, выручка нетто,
    EBIT, чистая прибыль, CF инвестора, накопленный CF, остаток кредита.
    """
    params = _load_params(project_name)
    if params is None:
        return json.dumps({"error": f"Проект не найден: {project_name}"}, ensure_ascii=False)

    rows = calculate(params)

    def _slim(r):
        return {
            "month":         r.month,
            "date":          r.date_str,
            "phase":         r.phase,
            "revenue_net":   round(r.revenue_net, 2),
            "ebit":          round(r.ebit, 2),
            "interest":      round(r.interest, 2),
            "net_profit":    round(r.net_profit, 2),
            "investor_cf":   round(r.investor_cf, 2),
            "cumulative_cf": round(r.cumulative_cf, 2),
            "loan_balance":  round(r.loan_balance, 2),
        }

    return json.dumps({
        "horizon_months":      params.horizon_months,
        "construction_months": params.construction_months,
        "first":               [_slim(r) for r in rows[:n_first]],
        "last":                [_slim(r) for r in rows[-n_last:]],
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def compare_projects(project_names: list[str]) -> str:
    """
    Сравнить KPI нескольких проектов одной таблицей.
    Возвращает массив {name, kpi} — без построения Excel.
    """
    out = []
    for name in project_names:
        params = _load_params(name)
        if params is None:
            out.append({"name": name, "error": "не найден"})
            continue
        rows = calculate(params)
        out.append({"name": name, "kpi": calculate_kpi(params, rows)})
    return json.dumps({"comparison": out}, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    mcp.run()
