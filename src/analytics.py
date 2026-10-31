"""Расчетный движок финансового анализа (Этап 2).

Показатели по разделам методологии (input/Methodology of analysis.md):
1. Ликвидность и платежеспособность
2. Деловая активность (коэффициент и дни)
3. Рентабельность и эффективность (ROS, ROA, ROE, EBITDA)
4. Структура капитала и долговая нагрузка
5. Горизонтальный анализ всех форм (BR-04)
6. Риски банкротства (Z-счет Альтмана — двух- и пятифакторная модели,
   модель Таффлера, анализ денежных потоков)

Данные для диаграмм и таблицы «Основные показатели ФХД» формируются
в млн руб. делением на 1000 (для проверки корректности) — FSA_brief.docx.
Отсутствующие данные -> «требует уточнения» (BR-02).
"""
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

MISSING = "требует уточнения"

# --- Нормативные диапазоны (reference/brief_rezult.pdf) --------------------
NORM_CURRENT = (1.5, 2.0)        # текущая ликвидность
NORM_QUICK = (0.8, 1.0)          # быстрая
NORM_CASH = (0.3, 0.5)           # абсолютная
NORM_OWN_SOURCES = (0.2, 0.5)    # обеспеченность собств. источниками
NORM_AUTONOMY = (0.5, None)      # фин. независимость >= 0,5
NORM_STABILITY = (0.6, 0.8)      # фин. устойчивость
NORM_LEVERAGE = (None, 1.0)      # финансовый рычаг / Debt-Equity <= 1,0

# Зоны Z-счета Альтмана (пятифакторная модель, reference/brief_rezult.pdf)
Z_GREEN, Z_RED = 2.9, 1.8
# Модель Таффлера
TAFFLER_GREEN, TAFFLER_RED = 0.3, 0.2

# Коды ключевых строк для таблицы «Основные показатели ФХД» (FSA_brief.docx)
KEY_LINES = [
    ("Баланс", "1100", "Внеоборотные активы"),
    ("Баланс", "1200", "Оборотные активы"),
    ("Баланс", "1300", "Собственный капитал"),
    ("Баланс", "1400", "Долгосрочные обязательства"),
    ("Баланс", "1500", "Краткосрочные обязательства"),
    ("ОФР", "2110", "Выручка"),
    ("ОФР", "2120", "Себестоимость"),
    ("ОФР", "2100", "Валовая прибыль (убыток)"),
    ("ОФР", "2210", "Коммерческие расходы"),
    ("ОФР", "2220", "Управленческие расходы"),
    ("ОФР", "2200", "Прибыль (убыток) от продаж"),
    ("ОФР", "2300", "Прибыль (убыток) до налогообложения"),
    ("ОФР", "2400", "Чистая прибыль (убыток)"),
    ("ОДДС", "4100", "Сальдо денежного потока от текущих операций"),
    ("ОДДС", "4400", "Сальдо денежного потока за период"),
]

BALANCE_ASSET_TOTALS = {"1100", "1200", "1600"}
BALANCE_LIAB_TOTALS = {"1300", "1400", "1500", "1600", "1700"}


# --- утилиты ---------------------------------------------------------------
def norm_code(value: Any) -> str:
    if value is None or (isinstance(value, float) and value != value):
        return ""
    s = str(value).strip().replace("\n", "").replace(" ", "")
    if s.lower() in ("", "nan", "none"):
        return ""
    m = re.match(r"^(\d{4,5})(?:\.0+)?$", s)
    return m.group(1) if m else s


def _num(value: Any) -> Optional[float]:
    try:
        if value is None or (isinstance(value, float) and value != value):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _sub(a, b) -> Optional[float]:
    return None if a is None or b is None else a - b


def _add(a, b) -> Optional[float]:
    return None if a is None or b is None else a + b


def _div(a, b) -> Optional[float]:
    if a is None or b is None or b == 0:
        return None
    return a / b


def _x100(x):
    return None if x is None else x * 100


def status_of(value: Optional[float],
              norm: Optional[Tuple[Optional[float], Optional[float]]]) -> str:
    """BR-01: статус по рекомендуемому диапазону (lo, hi)."""
    if norm is None:
        return "—" if value is not None else MISSING
    if value is None:
        return MISSING
    lo, hi = norm
    if (lo is not None and value < lo - 1e-9) or (hi is not None and value > hi + 1e-9):
        return "Риск"
    return "Безопасно"


def zzone_of(z: Optional[float]) -> str:
    if z is None:
        return MISSING
    if z > Z_GREEN:
        return "Безопасно"    # зеленая зона
    if z > Z_RED:
        return "Серая зона"   # зона неопределенности
    return "Риск"             # зона финансового риска


def taffler_zone(z: Optional[float]) -> str:
    if z is None:
        return MISSING
    if z > TAFFLER_GREEN:
        return "Безопасно"
    if z > TAFFLER_RED:
        return "Серая зона"
    return "Риск"


class AnalyticsEngine:
    """Расчет KPI, Z-счетов, горизонтального анализа и данных для диаграмм."""

    def __init__(self, forms: Dict[str, List[Dict[str, Any]]]):
        self.forms = forms
        self.years = self._extract_years()  # по убыванию: [2025, 2024, 2023]
        self.depreciation = None            # {год: амортизация} — задается в compute_all

    def _extract_years(self) -> List[int]:
        years = set()
        for rows in self.forms.values():
            for r in rows:
                for c in r:
                    m = re.match(r"^val_(\d{4})$", str(c))
                    if m:
                        years.add(int(m.group(1)))
        return sorted(years, reverse=True)

    # ------------------------- доступ к данным ---------------------------- #

    def v(self, form: str, code: str, year: int) -> Optional[float]:
        if year not in self.years:
            return None
        for r in self.forms.get(form, []):
            if norm_code(r.get("code")) == code:
                return _num(r.get(f"val_{year}"))
        return None

    def avg(self, form: str, code: str, year: int) -> Optional[float]:
        """Средняя величина периода; для крайнего года — значение на дату."""
        prev_years = [y for y in self.years if y < year]
        prev = self.v(form, code, prev_years[0]) if prev_years else None
        cur = self.v(form, code, year)
        if cur is None:
            return None
        return cur if prev is None else (cur + prev) / 2

    # --------------------------- главный метод ----------------------------- #

    def compute_all(self, depreciation: Optional[Dict[int, float]] = None) -> Dict[str, Any]:
        self.depreciation = depreciation
        return {
            "main_table": self.main_table(),           # млн руб.
            "kpi": self.kpi_merged(),                  # все показатели методологии
            "horizontal": self.horizontal_analysis(),  # dict: форма -> DataFrame
            "chart_data": self.chart_data(),           # dict: диаграмма -> DataFrame
        }

    # 1. Ликвидность и платежеспособность ------------------------------------
    def kpi_liquidity(self) -> List[Dict[str, Any]]:
        section = "Ликвидность и платежеспособность"
        for year in self.years:
            yield_year = year

            def val(code):
                return self.v("Баланс", code, yield_year)

            cur = _div(val("1200"), val("1500"))
            quick = _div(_sub(val("1200"), val("1210")), val("1500"))
            cash = _div(val("1250"), val("1500"))
            own_src = _div(_sub(val("1300"), val("1100")), val("1200"))
            autonomy = _div(val("1300"), val("1700"))
            yield [
                {"section": section, "name": "Коэффициент текущей ликвидности",
                 "formula": "стр.1200 / стр.1500", "values": {year: cur},
                 "range": "1,5–2,0", "norm": NORM_CURRENT,
                 "excel": "{B@1200}/{B@1500}"},
                {"section": section, "name": "Коэффициент быстрой ликвидности",
                 "formula": "(стр.1200 – стр.1210) / стр.1500",
                 "values": {year: quick}, "range": "0,8–1,0",
                 "norm": NORM_QUICK,
                 "excel": "({B@1200}-{B@1210})/{B@1500}"},
                {"section": section, "name": "Коэффициент абсолютной ликвидности",
                 "formula": "стр.1250 / стр.1500", "values": {year: cash},
                 "range": "0,3–0,5", "norm": NORM_CASH,
                 "excel": "{B@1250}/{B@1500}"},
                {"section": section,
                 "name": "Коэффициент обеспеченности собственными источниками",
                 "formula": "(стр.1300 – стр.1100) / стр.1200",
                 "values": {year: own_src}, "range": "0,2–0,5",
                 "norm": NORM_OWN_SOURCES,
                 "excel": "({B@1300}-{B@1100})/{B@1200}"},
                {"section": section, "name": "Коэффициент автономии",
                 "formula": "стр.1300 / стр.1700", "values": {year: autonomy},
                 "range": ">= 0,5", "norm": NORM_AUTONOMY,
                 "excel": "{B@1300}/{B@1700}"},
            ]

    # 2. Деловая активность ---------------------------------------------------
    def kpi_activity(self) -> List[Dict[str, Any]]:
        section = "Деловая активность (оборачиваемость)"
        for year in self.years:
            revenue = self.v("ОФР", "2110", year)
            cost = self.v("ОФР", "2120", year)
            cost_abs = abs(cost) if cost is not None else None

            debt_r = _div(revenue, self.avg("Баланс", "1230", year))
            cred_r = _div(cost_abs, self.avg("Баланс", "1520", year))
            wc = _sub(self.avg("Баланс", "1200", year),
                      self.avg("Баланс", "1500", year))
            wc_r = _div(revenue, wc)
            inv_r = _div(cost_abs, self.avg("Баланс", "1210", year))

            tdd = _div(365.0, debt_r)   # дебиторская, дней
            tpt = _div(365.0, cred_r)   # кредиторская, дней
            tst = _div(365.0, inv_r)    # запасы, дней
            twc = _div(365.0, wc_r)     # оборотный капитал, дней
            ccc = (_add(_add(tdd, tst), -tpt) if tpt is not None
                   else _add(tdd, tst))

            for name, coef, days, formula, excel_tpl in [
                ("Оборачиваемость дебиторской задолженности", debt_r, tdd,
                 "выручка (2110) / средняя стр.1230; дней=365/коэф.",
                 "{OF@2110}/{AVG@Баланс@1230}"),
                ("Оборачиваемость кредиторской задолженности", cred_r, tpt,
                 "себестоимость |2120| / средняя стр.1520; дней=365/коэф.",
                 "ABS({OF@2120})/{AVG@Баланс@1520}"),
                ("Оборачиваемость оборотного капитала", wc_r, twc,
                 "выручка (2110) / средний (стр.1200 – стр.1500); дней=365/коэф.",
                 "{OF@2110}/({AVG@Баланс@1200}-{AVG@Баланс@1500})"),
                ("Оборачиваемость запасов", inv_r, tst,
                 "себестоимость |2120| / средняя стр.1210; дней=365/коэф.",
                 "ABS({OF@2120})/{AVG@Баланс@1210}"),
            ]:
                yield [
                    {"section": section, "name": name + ", коэффициент",
                     "formula": formula, "values": {year: coef},
                     "range": "не установлен", "norm": None,
                     "excel": excel_tpl},
                    {"section": section, "name": name + ", дней",
                     "formula": formula, "values": {year: days},
                     "range": "не установлен", "norm": None,
                     "excel": "365/{K:%s}" % (name + ", коэффициент")},
                ]
            yield [
                {"section": section,
                 "name": "Цикл конверсии денежных средств, дней",
                 "formula": "Дни(ДЗ) + Дни(запасов) – Дни(КЗ)",
                 "values": {year: ccc},
                 "range": "не установлен", "norm": None,
                 "excel": "{K:Оборачиваемость дебиторской задолженности, дней}"
                          "+{K:Оборачиваемость запасов, дней}"
                          "-IFERROR({K:Оборачиваемость кредиторской "
                          "задолженности, дней},0)"},
            ]

    # 3. Рентабельность и эффективность ---------------------------------------
    def kpi_profitability(self) -> List[Dict[str, Any]]:
        section = "Рентабельность и эффективность"
        for year in self.years:
            revenue = self.v("ОФР", "2110", year)
            interest = self.v("ОФР", "2330", year)
            ebit = _add(self.v("ОФР", "2300", year),
                        abs(interest) if interest is not None else None)
            dep = (self.depreciation or {}).get(year)
            # EBITDA считается и без амортизации; если амортизация не введена,
            # в столбце диапазона ставится примечание «рассчитано без амортизации»
            ebitda = ebit if dep is None else _add(ebit, dep)
            top_down = self.v("ОФР", "2200", year) if dep is None \
                else _add(self.v("ОФР", "2200", year), dep)
            note = "рассчитано без амортизации" if dep is None else "—"

            dep_ref = "+{DEP}" if dep is not None else ""
            rows = [
                {"section": section, "name": "Рентабельность продаж (ROS), %",
                 "formula": "прибыль от продаж (2200) / выручка (2110) × 100",
                 "values": {year: _x100(_div(self.v("ОФР", "2200", year), revenue))},
                 "range": "не установлен", "norm": None,
                 "excel": "{OF@2200}/{OF@2110}*100"},
                {"section": section, "name": "Рентабельность активов (ROA), %",
                 "formula": "чистая прибыль (2400) / средняя величина активов (1600) × 100",
                 "values": {year: _x100(_div(self.v("ОФР", "2400", year),
                                             self.avg("Баланс", "1600", year)))},
                 "range": "не установлен", "norm": None,
                 "excel": "{OF@2400}/{AVG@Баланс@1600}*100"},
                {"section": section, "name": "Рентабельность собственного капитала (ROE), %",
                 "formula": "чистая прибыль (2400) / средняя величина СК (1300) × 100",
                 "values": {year: _x100(_div(self.v("ОФР", "2400", year),
                                             self.avg("Баланс", "1300", year)))},
                 "range": "не установлен", "norm": None,
                 "excel": "{OF@2400}/{AVG@Баланс@1300}*100"},
                {"section": section, "name": "EBITDA (снизу-вверх), тыс. руб.",
                 "formula": "2300 + |проценты к уплате (2330)| + амортизация",
                 "values": {year: ebitda}, "range": note, "norm": None,
                 "excel": "({OF@2300}+ABS({OF@2330})%s)" % dep_ref},
                {"section": section, "name": "EBITDA (сверху-вниз), тыс. руб.",
                 "formula": "2200 + амортизация",
                 "values": {year: top_down}, "range": note, "norm": None,
                 "excel": "({OF@2200}%s)" % dep_ref},
                {"section": section, "name": "Рентабельность по EBITDA, %",
                 "formula": "EBITDA (снизу-вверх) / выручка (2110) × 100",
                 "values": {year: _x100(_div(ebitda, revenue))},
                 "range": note, "norm": None,
                 "excel": "{K:EBITDA (снизу-вверх), тыс. руб.}/{OF@2110}*100"},
            ]
            yield rows

    # 4. Структура капитала и долговая нагрузка -------------------------------
    def kpi_capital(self) -> List[Dict[str, Any]]:
        section = "Структура капитала и долговая нагрузка"

        def ebit(year):
            interest = self.v("ОФР", "2330", year)
            return _add(self.v("ОФР", "2300", year),
                        abs(interest) if interest is not None else None)

        for year in self.years:

            def val(code):
                return self.v("Баланс", code, year)

            d_e = _div(_add(val("1400"), val("1500")), val("1300"))
            interest = self.v("ОФР", "2330", year)
            icr = _div(ebit(year), abs(interest)) if interest is not None else None
            wc = _sub(val("1200"), val("1500"))
            stability = _div(_add(val("1300"), val("1400")), val("1700"))
            yield [
                {"section": section, "name": "Debt/Equity",
                 "formula": "(стр.1400 + стр.1500) / стр.1300",
                 "values": {year: d_e}, "range": "<= 1,0",
                 "norm": NORM_LEVERAGE,
                 "excel": "({B@1400}+{B@1500})/{B@1300}"},
                {"section": section,
                 "name": "Interest Coverage Ratio (покрытие процентов)",
                 "formula": "(2300 + |проценты к уплате (2330)|) / |проценты к уплате (2330)|",
                 "values": {year: icr},
                 "range": "не установлен", "norm": None,
                 "excel": "({OF@2300}+ABS({OF@2330}))/ABS({OF@2330})"},
                {"section": section, "name": "Working capital, тыс. руб.",
                 "formula": "стр.1200 – стр.1500",
                 "values": {year: wc}, "range": "—", "norm": None,
                 "excel": "{B@1200}-{B@1500}"},
                {"section": section, "name": "Коэффициент финансовой устойчивости",
                 "formula": "(стр.1300 + стр.1400) / стр.1700",
                 "values": {year: stability},
                 "range": "0,6–0,8", "norm": NORM_STABILITY,
                 "excel": "({B@1300}+{B@1400})/{B@1700}"},
                {"section": section, "name": "Коэффициент финансового рычага",
                 "formula": "(стр.1400 + стр.1500) / стр.1300",
                 "values": {year: d_e}, "range": "<= 1,0",
                 "norm": NORM_LEVERAGE,
                 "excel": "({B@1400}+{B@1500})/{B@1300}"},
            ]

    # 6. Риски банкротства ------------------------------------------------------
    def kpi_bankruptcy(self) -> List[Dict[str, Any]]:
        section = "Риски банкротства и потери платежеспособности"
        for year in self.years:

            def val(code):
                return self.v("Баланс", code, year)

            ta = val("1600")
            x1 = _div(_sub(val("1200"), val("1500")), ta)            # (ОА–КО)/Активы
            x2 = _div(val("1370"), ta)                               # НП/Активы
            interest = self.v("ОФР", "2330", year)
            ebit = _add(self.v("ОФР", "2300", year),
                        abs(interest) if interest is not None else None)
            x3 = _div(ebit, ta)                                       # EBIT/Активы
            x4 = _div(val("1300"), _add(val("1400"), val("1500")))    # СК/ЗС
            x5 = _div(self.v("ОФР", "2110", year), ta)                # Выручка/Активы

            def weighted(*pairs, offset=0.0):
                if any(x is None for _, x in pairs):
                    return None
                return offset + sum(k * x for k, x in pairs)

            z5 = weighted((1.2, x1), (1.4, x2), (3.3, x3),
                          (0.6, x4), (1.0, x5))
            cr = _div(val("1200"), val("1500"))   # текущая ликвидность
            debt_share = _div(_add(val("1400"), val("1500")), val("1700"))
            z2 = weighted((-1.073, cr), (0.0579, debt_share), offset=-0.3877)
            taf = weighted(
                (0.13, _div(self.v("ОФР", "2300", year), val("1500"))),
                (0.13, _div(val("1200"), _add(val("1400"), val("1500")))),
                (0.13, _div(val("1500"), ta)),
                (0.13, _div(self.v("ОФР", "2110", year), ta)),
                offset=0.53)
            cf_to_debt = _div(self.v("ОДДС", "4100", year),
                              _add(val("1400"), val("1500")))
            profit_quality = _div(self.v("ОДДС", "4100", year),
                                  self.v("ОФР", "2400", year))
            yield [
                {"section": section,
                 "name": "Z-счет Альтмана (двухфакторная модель)",
                 "formula": "Z = –0,3877 – 1,073·Ктек.ликв. + 0,0579·Доля заем. средств; "
                            "Z<0 — вероятность банкротства < 50%",
                 "values": {year: z2}, "range": "< 0", "norm": (None, 0.0),
                 "excel": "-0.3877-1.073*{K:Коэффициент текущей ликвидности}"
                          "+0.0579*({B@1400}+{B@1500})/{B@1700}"},
                {"section": section,
                 "name": "Z-счет Альтмана (пятифакторная модель)",
                 "formula": "Z = 1,2·X1 + 1,4·X2 + 3,3·X3 + 0,6·X4 + 1,0·X5",
                 "values": {year: z5},
                 "range": "> 2,9 — зеленая; 1,8–2,9 — серая; < 1,8 — красная",
                 "norm": None, "zone_fn": zzone_of,
                 "excel": "1.2*({B@1200}-{B@1500})/{B@1600}"
                          "+1.4*{B@1370}/{B@1600}"
                          "+3.3*({OF@2300}+ABS({OF@2330}))/{B@1600}"
                          "+0.6*{B@1300}/({B@1400}+{B@1500})"
                          "+{OF@2110}/{B@1600}"},
                {"section": section, "name": "Модель Таффлера",
                 "formula": "Z = 0,53 + 0,13·(X1 + X2 + X3 + X4)",
                 "values": {year: taf},
                 "range": "> 0,3 — устойчив; 0,2–0,3 — серая; < 0,2 — риск",
                 "norm": None, "zone_fn": taffler_zone,
                 "excel": "0.53+0.13*({OF@2300}/{B@1500}"
                          "+{B@1200}/({B@1400}+{B@1500})"
                          "+{B@1500}/{B@1600}+{OF@2110}/{B@1600})"},
                {"section": section,
                 "name": "Операционный денежный поток / обязательства",
                 "formula": "стр.4100 (ОДДС) / (стр.1400 + стр.1500)",
                 "values": {year: cf_to_debt},
                 "range": "не установлен", "norm": None,
                 "excel": "{OD@4100}/({B@1400}+{B@1500})"},
                {"section": section,
                 "name": "Операционный денежный поток / чистая прибыль "
                         "(проверка «бумажной» прибыли)",
                 "formula": "стр.4100 (ОДДС) / стр.2400 (ОФР)",
                 "values": {year: profit_quality},
                 "range": "не установлен", "norm": None,
                 "excel": "{OD@4100}/{OF@2400}"},
            ]

    # --------------------------- сводка KPI ------------------------------- #
    def kpi_merged(self) -> List[Dict[str, Any]]:
        """Слияние по разделам/показателям за все годы; статус — по отчетному году."""
        merged: Dict[tuple, Dict[str, Any]] = {}
        order = []
        generators = (self.kpi_liquidity(), self.kpi_activity(),
                      self.kpi_profitability(), self.kpi_capital(),
                      self.kpi_bankruptcy())
        for gen in generators:
            for batch in gen:
                for r in batch:
                    key = (r["section"], r["name"])
                    if key not in merged:
                        merged[key] = {"section": r["section"], "name": r["name"],
                                       "formula": r["formula"], "range": r["range"],
                                       "norm": r.get("norm"),
                                       "zone_fn": r.get("zone_fn"),
                                       "missing": r.get("missing"),
                                       "excel": r.get("excel"),
                                       "values": {}}
                        order.append(key)
                    merged[key]["values"].update(r["values"])

        out = []
        for key in order:
            e = merged[key]
            vals = e["values"]
            e["values"] = {y: (None if vals.get(y) is None else round(vals[y], 2))
                           for y in self.years}
            # Статус/зона — по значению отчетного (первого) года, где данные есть
            latest = None
            for y in self.years:
                if vals.get(y) is not None:
                    latest = vals[y]
                    break
            if e.get("zone_fn") is not None:
                e["status"] = e["zone_fn"](latest)
            else:
                e["status"] = status_of(latest, e.get("norm"))
            out.append(e)
        self._kpi_list = out
        return out

    # ------------------ Таблица основных показателей ----------------------- #
    def main_table(self) -> pd.DataFrame:
        """«Основные показатели финансово-хозяйственной деятельности», млн руб."""
        if not self.years:
            return pd.DataFrame()
        data: Dict[str, list] = {"Показатель": [], "Источник": []}
        for year in self.years:
            data[f"{year} г., млн руб."] = []
        for form, code, name in KEY_LINES:
            data["Показатель"].append(name)
            data["Источник"].append(f"{form}, стр. {code}")
            for year in self.years:
                val = self.v(form, code, year)
                data[f"{year} г., млн руб."].append(
                    MISSING if val is None else round(val / 1000, 1))
        return pd.DataFrame(data)

    # ------------------- Горизонтальный анализ (BR-04) --------------------- #
    def horizontal_analysis(self) -> Dict[str, pd.DataFrame]:
        out: Dict[str, pd.DataFrame] = {}
        pairs = [(self.years[i], self.years[i + 1])
                 for i in range(len(self.years) - 1)]
        if pairs:
            out["Баланс"] = self._horizontal_df(
                "Баланс", pairs, skip=BALANCE_ASSET_TOTALS)
            out["ОФР"] = self._horizontal_df("ОФР", pairs)
            out["ОДДС"] = self._horizontal_df("ОДДС", pairs)
        oik = self._horizontal_oik(pairs)
        if oik is not None:
            out["ОИК"] = oik
        return out

    def _horizontal_df(self, form: str, pairs,
                       skip: Optional[set] = None) -> pd.DataFrame:
        rows = [r for r in self.forms.get(form, []) if norm_code(r.get("code"))]
        data = []
        for r in rows:
            code = norm_code(r.get("code"))
            if skip and code in skip:
                continue
            vals = {int(c[4:]): _num(v)
                    for c, v in r.items()
                    if str(c).startswith("val_")}
            entry = {"Код": code, "Показатель": str(r.get("indicator", "")).strip()}
            for cur, prev in pairs:
                a, b = vals.get(cur), vals.get(prev)
                if a is None or b is None:
                    entry[f"{cur} к {prev}"] = MISSING
                elif b == 0:
                    entry[f"{cur} к {prev}"] = (
                        f"{(a - b) / 1000:+,.1f} млн (—)".replace(",", " ")
                        if a != b else "0")
                else:
                    pct = (a - b) / abs(b) * 100
                    entry[f"{cur} к {prev}"] = (f"{(a - b) / 1000:+,.1f} млн "
                                                f"({pct:+.1f}%)").replace(",", " ")
            data.append(entry)
        cols = ["Код", "Показатель"] + [f"{c} к {p}" for c, p in pairs]
        return pd.DataFrame(data, columns=cols)

    def _horizontal_oik(self, pairs) -> Optional[pd.DataFrame]:
        """Для ОИК — динамика капитала на отчетные даты (столбец «Итого»)."""
        dates = {2023: ("33100", "3100"), 2024: ("63200", "3200"),
                 2025: ("83300", "3300")}
        rec = []
        for y in self.years:
            row = next((r for r in self.forms.get("ОИК", [])
                        if norm_code(r.get("code")) in dates.get(y, ())), None)
            rec.append({"year": y,
                        "total": _num(row.get("Итого")) if row else None})
        rec = [r for r in rec if r["total"] is not None]
        if len(rec) < 2:
            return None
        rows = []
        for a, b in zip(rec, rec[1:]):
            # a — более поздний год, b — предыдущий
            change = a["total"] - b["total"]
            rows.append({"Период": f"{a['year']} к {b['year']}",
                         "Капитал на конец периода, млн руб.": round(a["total"] / 1000, 1),
                         "Изменение": (f"{change / 1000:+,.1f} млн "
                                       f"({change / abs(b['total']) * 100:+.1f}%)"
                                       if b["total"] else "—").replace(",", " ")})
        return pd.DataFrame(rows)

    # ------------------- Данные для диаграмм (млн руб.) -------------------- #
    def chart_data(self) -> Dict[str, pd.DataFrame]:
        if not self.years:
            return {}
        year = self.years[0]
        prev_year = self.years[1] if len(self.years) > 1 else None
        out: Dict[str, pd.DataFrame] = {}
        col = f"Значение, млн руб. ({year})"

        # «Структура активов» — все строки актива с кодом, без промежуточных итогов
        rec = []
        for r in self.forms.get("Баланс", []):
            code = norm_code(r.get("code"))
            if not code or code in BALANCE_ASSET_TOTALS:
                continue
            if not 1110 <= int(code) <= 1269:
                continue
            val = self.v("Баланс", code, year)
            rec.append({"Код": code, "Статья": str(r.get("indicator", "")).strip(),
                        col: MISSING if val is None else round(val / 1000, 1)})
        if rec:
            out["Структура активов"] = pd.DataFrame(rec)

        # «Структура капитала» — все строки пассива с кодом, без промежуточных итогов
        rec = []
        for r in self.forms.get("Баланс", []):
            code = norm_code(r.get("code"))
            if not code or code in BALANCE_LIAB_TOTALS:
                continue
            if not 1310 <= int(code) <= 1549:
                continue
            val = self.v("Баланс", code, year)
            rec.append({"Код": code, "Статья": str(r.get("indicator", "")).strip(),
                        col: MISSING if val is None else round(val / 1000, 1)})
        if rec:
            out["Структура капитала"] = pd.DataFrame(rec)

        # «Динамика изменений статей Баланса» (включая стр. 1700), без
        # промежуточных итогов разделов («Итого по разделу I … V», 1600)
        if prev_year:
            rec = []
            for r in self.forms.get("Баланс", []):
                code = norm_code(r.get("code"))
                name_txt = str(r.get("indicator", "")).strip().lower()
                if not code or code == "1600" or name_txt.startswith("итого по разделу"):
                    continue
                a = self.v("Баланс", code, year)
                b = self.v("Баланс", code, prev_year)
                change = round((a - b) / 1000, 1) if a is not None and b is not None else None
                pct = (a - b) / abs(b) * 100 if change is not None and b else None
                rec.append({"Код": code,
                            "Статья": str(r.get("indicator", "")).strip(),
                            "Изменение, млн руб.": MISSING if change is None
                            else change,
                            "Изменение, %": MISSING if pct is None
                            else f"{pct:+.1f}"})
            if rec:
                out["Динамика изменений статей Баланса"] = pd.DataFrame(rec)

        # «Финансовые результаты» — все строки ОФР с кодом до 2400 включительно,
        # отчетный и предыдущий год (для линейчатой диаграммы с группировкой)
        rec = []
        for r in self.forms.get("ОФР", []):
            code = norm_code(r.get("code"))
            if not code or int(code) > 2400:
                continue
            val = self.v("ОФР", code, year)
            val_prev = self.v("ОФР", code, prev_year) if prev_year else None
            rec.append({
                "Код": code,
                "Статья": str(r.get("indicator", "")).strip(),
                col: MISSING if val is None else round(val / 1000, 1),
                **({f"Значение, млн руб. ({prev_year})":
                       MISSING if val_prev is None else round(val_prev / 1000, 1)}
                   if prev_year else {}),
            })
        if rec:
            out["Финансовые результаты"] = pd.DataFrame(rec)

        return out