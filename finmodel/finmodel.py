# -*- coding: utf-8 -*-
"""
Финансовая модель: помесячный расчёт денежных потоков склада.

Алгоритм:
  Фаза строительства (мес. 1..T_c):
    - CAPEX равномерно: construction_cost / T_c в месяц
    - Кредит траншами: loan_amount / T_c в месяц
    - Инвестор платит: долю СС + проценты на остаток кредита

  Фаза эксплуатации (мес. T_c+1..T_h):
    - Выручка = area × rental_rate × (1 + inflation)^year_ops
    - НДС входной из строительства возмещается в первые 3 мес.
    - Аннуитетное погашение кредита
    - Налоги: прибыль 20%, имущество 2,2%/год, НДС 22%
    - Амортизация: линейная, 40 лет (480 мес.)

  CF инвестора = Чистая прибыль + Амортизация − Погашение кредита − НДС к уплате
"""
from __future__ import annotations
import json
from dataclasses import dataclass, field
from typing import List, Optional
from pathlib import Path


# ─── ПАРАМЕТРЫ ПРОЕКТА ───────────────────────────────────────────────────────

@dataclass
class ProjectParams:
    name: str
    start_date: str              # "YYYY-MM"
    area_sqm: float              # кв.м.
    construction_cost: float     # руб. с НДС
    loan_share: float            # доля кредита 0..1
    cbr_rate: float              # ставка ЦБ
    loan_rate: float             # кредитная ставка (годовая)
    rental_rate_sqm: float       # руб./кв.м./мес. с НДС
    inflation_rate: float        # годовая инфляция
    construction_months: int     # срок строительства, мес.
    horizon_months: int          # горизонт планирования, мес.
    vat_rate: float = 0.22
    income_tax_rate: float = 0.20
    property_tax_rate: float = 0.022
    operating_cost_pct: float = 0.15   # % от выручки без НДС
    useful_life_years: int = 40

    @classmethod
    def from_json(cls, path: str | Path) -> "ProjectParams":
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        fields = cls.__dataclass_fields__
        return cls(**{k: v for k, v in d.items() if k in fields})

    @property
    def cost_ex_vat(self) -> float:
        return self.construction_cost / (1 + self.vat_rate)

    @property
    def vat_construction(self) -> float:
        return self.construction_cost - self.cost_ex_vat

    @property
    def loan_amount(self) -> float:
        return self.construction_cost * self.loan_share

    @property
    def equity_amount(self) -> float:
        return self.construction_cost * (1 - self.loan_share)

    @property
    def ops_months(self) -> int:
        return self.horizon_months - self.construction_months


# ─── СТРОКА ДЕНЕЖНЫХ ПОТОКОВ ─────────────────────────────────────────────────

@dataclass
class MonthRow:
    month: int
    date_str: str
    phase: str
    # Выручка
    revenue_gross: float = 0.0
    revenue_net: float = 0.0
    vat_output: float = 0.0
    vat_input: float = 0.0
    vat_payable: float = 0.0
    # P&L
    operating_costs: float = 0.0
    depreciation: float = 0.0
    property_tax: float = 0.0
    ebit: float = 0.0
    interest: float = 0.0
    ebt: float = 0.0
    income_tax: float = 0.0
    net_profit: float = 0.0
    # Движение по балансу
    capex: float = 0.0
    loan_drawdown: float = 0.0
    principal_paid: float = 0.0
    loan_balance: float = 0.0
    # НДС
    vat_refund: float = 0.0    # Возврат НДС из налоговой (позитивный = приток)
    # CF инвестора
    investor_cf: float = 0.0
    cumulative_cf: float = 0.0


# ─── РАСЧЁТНЫЙ ДВИЖОК ────────────────────────────────────────────────────────

def _date_str(start_year: int, start_month: int, offset: int) -> str:
    """Вернуть строку ДД.ММ.ГГГГ для (start + offset) месяцев. День всегда 01."""
    total = start_month + offset - 1
    y = start_year + total // 12
    m = total % 12 + 1
    return f"01.{m:02d}.{y:04d}"


def calculate(params: ProjectParams) -> List[MonthRow]:
    """Рассчитать помесячные денежные потоки."""
    p = params
    sy, sm = map(int, p.start_date.split("-"))

    # Предварительные расчёты
    capex_monthly = p.construction_cost / p.construction_months
    loan_drawdown_m = p.loan_amount / p.construction_months
    equity_monthly = p.equity_amount / p.construction_months
    depreciation_m = p.cost_ex_vat / (p.useful_life_years * 12)

    # Аннуитетный платёж в фазе эксплуатации
    r_m = p.loan_rate / 12
    n = p.ops_months
    if r_m > 0 and n > 0:
        annuity = p.loan_amount * r_m / (1 - (1 + r_m) ** (-n))
    else:
        annuity = p.loan_amount / n if n > 0 else 0.0

    # НДС с строительства возмещается в первые 3 мес. эксплуатации
    vat_credit_per_month = p.vat_construction / 3

    loan_balance = 0.0
    book_value = p.cost_ex_vat
    cumulative_cf = 0.0
    rows: List[MonthRow] = []

    for t in range(1, p.horizon_months + 1):
        row = MonthRow(
            month=t,
            date_str=_date_str(sy, sm, t),
            phase="Строительство" if t <= p.construction_months else "Эксплуатация",
        )

        if t <= p.construction_months:
            # ── Фаза строительства ──────────────────────────────────────────
            loan_balance += loan_drawdown_m
            row.capex = capex_monthly
            row.loan_drawdown = loan_drawdown_m
            row.interest = loan_balance * p.loan_rate / 12
            row.loan_balance = loan_balance
            row.net_profit = -row.interest
            # CF инвестора: выплачивает долю СС + проценты на накопленный кредит
            row.investor_cf = -(equity_monthly + row.interest)

        else:
            # ── Фаза эксплуатации ───────────────────────────────────────────
            ops_t = t - p.construction_months  # 1-й месяц = 1

            # Ежегодная индексация
            years_done = (ops_t - 1) // 12
            infl_factor = (1 + p.inflation_rate) ** years_done

            # Выручка
            row.revenue_gross = p.area_sqm * p.rental_rate_sqm * infl_factor
            row.revenue_net = row.revenue_gross / (1 + p.vat_rate)
            row.vat_output = row.revenue_gross - row.revenue_net

            # НДС входной (возмещение из строительства в первые 3 мес.)
            if ops_t <= 3:
                row.vat_input = vat_credit_per_month
            # Чистый НДС-поток: положительный = платим, отрицательный = получаем возврат
            _net_vat = row.vat_output - row.vat_input
            row.vat_payable = max(0.0, _net_vat)
            row.vat_refund = max(0.0, -_net_vat)

            # Операционные расходы (без НДС)
            row.operating_costs = row.revenue_net * p.operating_cost_pct

            # Амортизация и налог на имущество
            row.depreciation = depreciation_m
            row.property_tax = max(0.0, book_value) * p.property_tax_rate / 12
            book_value -= depreciation_m

            # P&L
            row.ebit = (row.revenue_net
                        - row.operating_costs
                        - row.depreciation
                        - row.property_tax)
            row.interest = loan_balance * p.loan_rate / 12
            row.ebt = row.ebit - row.interest
            row.income_tax = max(0.0, row.ebt * p.income_tax_rate)
            row.net_profit = row.ebt - row.income_tax

            # Аннуитетное погашение кредита
            principal = min(max(0.0, annuity - row.interest), loan_balance)
            loan_balance = max(0.0, loan_balance - principal)
            row.principal_paid = principal
            row.loan_balance = loan_balance

            # CF инвестора (свободный денежный поток на СС)
            # _net_vat: положительный = выплата НДС (отток), отрицательный = возврат (приток)
            row.investor_cf = (row.net_profit
                               + row.depreciation
                               - row.principal_paid
                               - _net_vat)

        cumulative_cf += row.investor_cf
        row.cumulative_cf = cumulative_cf
        rows.append(row)

    return rows


# ─── KPI ─────────────────────────────────────────────────────────────────────

def _npv_m(rate_m: float, cfs: list) -> float:
    return sum(cf / (1 + rate_m) ** t for t, cf in enumerate(cfs))


def _irr_annual(cfs: list) -> Optional[float]:
    """Годовой IRR через бисекцию по месячному IRR."""
    if not any(x < 0 for x in cfs) or not any(x > 0 for x in cfs):
        return None
    lo, hi = -0.9999 / 12, 5.0 / 12
    # Убедимся, что знаки разные
    f_lo = _npv_m(lo, cfs)
    f_hi = _npv_m(hi, cfs)
    if f_lo * f_hi > 0:
        # Попробуем расширить диапазон
        hi = 50.0 / 12
        f_hi = _npv_m(hi, cfs)
        if f_lo * f_hi > 0:
            return None
    for _ in range(600):
        mid = (lo + hi) / 2
        f_mid = _npv_m(mid, cfs)
        if abs(f_mid) < 1e-4:
            break
        if f_lo * f_mid < 0:
            hi = mid
            f_hi = f_mid
        else:
            lo = mid
            f_lo = f_mid
    return (1 + mid) ** 12 - 1


def calculate_kpi(params: ProjectParams, rows: List[MonthRow]) -> dict:
    """Рассчитать KPI проекта."""
    p = params

    # Серия CF для NPV/IRR: мес. 0 = вложение СС (единовременно), мес. 1..T = monthly CF
    # Мы распределяем equity по месяцам строительства, поэтому:
    # cf_series[0] = -equity (условно в момент t=0)
    # cf_series[1..T] = investor_cf из rows
    cf_series = [-p.equity_amount] + [r.investor_cf for r in rows]

    npv_val = _npv_m(p.cbr_rate / 12, cf_series)
    irr_val = _irr_annual(cf_series)

    # Срок окупаемости (с учётом вложения СС в момент 0)
    cumul = -p.equity_amount
    payback_month = None
    for r in rows:
        cumul += r.investor_cf
        if cumul >= 0 and payback_month is None:
            payback_month = r.month

    # ROI = чистая прибыль за горизонт / вложенные СС
    total_net_profit = sum(r.net_profit for r in rows)
    roi = total_net_profit / p.equity_amount if p.equity_amount > 0 else 0.0

    return {
        "npv": round(npv_val, 2),
        "irr_annual": round(irr_val * 100, 2) if irr_val is not None else None,
        "roi": round(roi * 100, 2),
        "payback_months": payback_month,
        "total_net_profit": round(total_net_profit, 2),
        "total_income_tax": round(sum(r.income_tax for r in rows), 2),
        "total_vat_paid": round(sum(r.vat_payable for r in rows), 2),
        "total_vat_refund": round(sum(r.vat_refund for r in rows), 2),
        "total_property_tax": round(sum(r.property_tax for r in rows), 2),
        "total_interest_paid": round(sum(r.interest for r in rows), 2),
        "total_revenue_net": round(sum(r.revenue_net for r in rows), 2),
        "total_revenue_gross": round(sum(r.revenue_gross for r in rows), 2),
        "total_operating_costs": round(sum(r.operating_costs for r in rows), 2),
    }
