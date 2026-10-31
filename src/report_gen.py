"""Генератор анонимизированных отчетов (Этап 4).

Создает пакет из трех файлов (FR-06…FR-09, BR-03):
1. explonation.xlsx — листы «Исходные данные», «Расчеты» и «Горизонтальный
   анализ» (каждое значение — Excel-формула в отдельной ячейке со ссылками
   на «Исходные данные», чтобы расчет можно было проверить), «Данные для
   диаграмм» (значения тоже формулами) и «Диаграммы» — нативные
   редактируемые диаграммы Excel, построенные на «Данных для диаграмм».
2. rezult.docx — заключение по тексту reference/brief_rezult.pdf (пункты 1,
   2, 3, 4, концовка); выводы по 6 разделам методологии; п. 3.7 — влияние
   изменения статей Пассива на статьи Актива; таблица «Риск банкротства»;
   Приложение 1 — диаграммы, скопированные из explonation.xlsx (OLE-объекты,
   редактируемые; при недоступном Excel — PNG-графики).
3. presentation.pptx — нативные (редактируемые) диаграммы PowerPoint,
   вывод на каждом слайде, финальный слайд в 2 колонки «Выводы» /
   «Рекомендации» с конкретными числовыми действиями.

Диаграммы docx (fallback PNG) рендерятся matplotlib в палитре
charts.BASE_STYLE, валидированной скиллом dataviz (Этап 3).

Анонимизация (BR-03): строки с ИНН (10/12 цифр), «ООО/АО/ИП…», ФИО,
должностями, адресами, телефонами заменяются на «[ДАННЫЕ УДАЛЕНЫ]».
После генерации каждого файла выполняется контрольная проверка текста
(verify_anonymity) — результат выводится в интерфейсе.

Отсутствующие данные -> «требует уточнения» (BR-02).
"""
import os
import re
import shutil
import tempfile
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from analytics import (AnalyticsEngine, KEY_LINES, MISSING, zzone_of,
                       taffler_zone)
from charts import (SMALL_LABEL, accent_colors, merge_small)  # круговые: «Прочие <1%», акценты топ-3

ANON_MARK = "[ДАННЫЕ УДАЛЕНЫ]"

# --- Анонимизация (BR-03): только текстовые значения ------------------------
_INN_RE = re.compile(r"(?<!\d)\d{10}(?!\d)|(?<!\d)\d{12}(?!\d)")
_ORG_RE = re.compile(r"(?<![А-ЯЁа-яёA-Za-z])(ООО|ОАО|ЗАО|ПАО|АО|ИП|ГУП|"
                     r"ФГУП|МУП|НИИ)(?![А-ЯЁа-яёA-Za-z])")
_FIO_RE = re.compile(r"[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.")
_KEYWORDS = ("фио", "должность", "инн", "огрн", "кпп", "персональные данные")
_PHONE_RE = re.compile(r"(?<!\d)(\+7|8)[\s(]?\d{3}[\)\s-]?\d{3}[\s-]?\d{2}"
                       r"[\s-]?\d{2}")
_ADDR_RE = re.compile(r"(ул\.|просп\.|проспект,?\s|шоссе\s|пер\.\s|р-н\s)",
                      re.IGNORECASE)

# Порядок диаграмм = порядок их создания на листе «Диаграммы» и в Прилож. 1
CHART_ORDER = ["Структура активов", "Структура капитала",
               "Динамика изменений статей Баланса", "Финансовые результаты",
               "Ликвидность", "Финансовая устойчивость", "Рентабельность",
               "Оборачиваемость, дней"]

# Дефолтные палитры (если пользовательский стиль не задан)
_DEFAULT_RAMPS = ["66B57F", "3E9E5C", "2B8849", "1E7139", "155C2C",
                  "0E4821", "AAA89F", "89877F", "6B6963", "4F4E48",
                  "383733", "242321"]
_DEFAULT_SERIES = ["2A78D6", "EB6834", "8A877F"]
_DEFAULT_DYNAMICS = "359C56"

def _style_ramps(style) -> List[str]:
    ramps = (style or {}).get("green_ramp") or []
    ramps += (style or {}).get("gray_ramp") or []
    if not ramps:
        ramps = list(_DEFAULT_RAMPS)
    return [c.lstrip("#") for c in ramps]

# Группы KPI для графиков «коэффициенты/дни по годам» (единые с app.py)
KPI_GROUPS = {
    "Ликвидность": {
        "Коэффициент обеспеченности собственными источниками": "Обеспеч. СОС",
        "Коэффициент текущей ликвидности": "Текущая",
        "Коэффициент быстрой ликвидности": "Быстрая",
        "Коэффициент абсолютной ликвидности": "Абсолютная",
    },
    "Финансовая устойчивость": {
        "Коэффициент автономии": "Независимость (>=0,5)",
        "Коэффициент финансовой устойчивости": "Устойчивость",
        "Коэффициент финансового рычага": "Рычаг",
    },
    "Рентабельность": {
        "Рентабельность продаж (ROS), %": "ROS",
        "Рентабельность активов (ROA), %": "ROA",
        "Рентабельность собственного капитала (ROE), %": "ROE",
        "Рентабельность по EBITDA, %": "EBITDA",
    },
    "Оборачиваемость, дней": {
        "Оборачиваемость дебиторской задолженности, дней": "ДЗ",
        "Оборачиваемость кредиторской задолженности, дней": "КЗ",
        "Оборачиваемость запасов, дней": "Запасы",
        "Оборачиваемость оборотного капитала, дней": "Оборотн. кап.",
        "Цикл конверсии денежных средств, дней": "ЦКДС",
    },
}
KPI_KIND = {"Ликвидность": (2,), "Финансовая устойчивость": (2,),
            "Рентабельность": (1,), "Оборачиваемость, дней": (0,)}


# --- Форматирование ----------------------------------------------------------
def ru_num(v: Optional[float], nd: int = 1) -> str:
    """Число в русском формате (пробел — тысячи, запятая — дробная часть)."""
    if v is None:
        return MISSING
    s = f"{v:,.{nd}f}"
    return s.replace(",", " ").replace(".", ",")


def status_word(status: str) -> str:
    return status if status not in ("", None) else "—"


# --- Анонимизация (BR-03) ----------------------------------------------------
def is_private(text: Any) -> bool:
    """True, если строка похоже содержит персональные данные (BR-03)."""
    if not isinstance(text, str) or not text.strip():
        return False
    low = text.lower()
    return bool(
        _INN_RE.search(text) or _ORG_RE.search(text) or _FIO_RE.search(text)
        or _PHONE_RE.search(text) or _ADDR_RE.search(text)
        or any(kw in low for kw in _KEYWORDS)
    )


def anonymize(text: Any) -> Any:
    """Замена строки с персональными данными на «[ДАННЫЕ УДАЛЕНЫ]» (BR-03)."""
    if isinstance(text, str) and is_private(text):
        return ANON_MARK
    return text


def verify_anonymity(path: str) -> List[str]:
    """Контрольная проверка готового файла; список найденных нарушений."""
    violations = []

    def check(cell, where):
        # Числовые значения и их строковые представления — не проверяются,
        # чтобы не ловить 10–12-значные суммы отчетности вместо ИНН
        if isinstance(cell, (int, float)) or cell is None:
            return
        text = str(cell)
        if not text.strip():
            return
        if re.fullmatch(r"[\s\d.,()+\-–—°%«»\"]*", text):
            return
        if is_private(text):
            violations.append(f"{where}: «{text.strip()[:60]}»")

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        check(cell, f"xlsx, лист «{ws.title}»")
            wb.close()
        elif ext == ".docx":
            import docx
            doc = docx.Document(path)
            for p in doc.paragraphs:
                check(p.text, "docx, абзац")
            for t in doc.tables:
                for row in t.rows:
                    for c in row.cells:
                        check(c.text, "docx, таблица")
        elif ext == ".pptx":
            from pptx import Presentation
            prs = Presentation(path)
            for i, slide in enumerate(prs.slides, 1):
                for shape in slide.shapes:
                    if getattr(shape, "has_text_frame", False):
                        check(shape.text_frame.text, f"pptx, слайд {i}")
    except Exception as e:  # noqa: BLE001 — проверка не должна ломать выдачу
        return [f"{path}: не удалось выполнить проверку ({e})"]
    return violations


# --- Извлечение чисел --------------------------------------------------------
def _num(v: Any) -> Optional[float]:
    try:
        if v is None or (isinstance(v, float) and v != v):
            return None
        if isinstance(v, str):
            return None if not v.strip() else float(v.replace(" ", "")
                                                     .replace(",", "."))
        return float(v)
    except (TypeError, ValueError):
        return None


def _ser_no_invert(ser) -> None:
    """<c:invertIfNegative val="0"/>: python-pptx не открывает этот узел,
    а без него Excel/PowerPoint рисуют отрицательные столбцы белыми."""
    from pptx.oxml.ns import qn
    el = ser._element
    if el.find(qn("c:invertIfNegative")) is None:
        inv = el.makeelement(qn("c:invertIfNegative"), {"val": "0"})
        prev = el.find(qn("c:spPr"))
        if prev is None:
            prev = el.find(qn("c:tx"))
        # порядок схемы: idx, order, tx, spPr, invertIfNegative, dPt, dLbls…
        prev.addnext(inv)


class ChartRenderer:
    """PNG-диаграммы (fallback для docx, если Excel COM недоступен;
    палитра charts.BASE_STYLE)."""

    def __init__(self, results: Dict[str, Any], style: Dict[str, Any],
                 out_dir: str):
        self.results = results
        self.style = style
        self.out_dir = out_dir
        self._imgs: Dict[str, str] = {}
        plt.rcParams.update({
            "font.family": "DejaVu Sans", "text.color": style["ink"],
            "figure.dpi": 200,
        })

    def _strip_axes(self, ax):
        for side in ("top", "right", "left", "bottom"):
            ax.spines[side].set_visible(False)
        ax.tick_params(colors=self.style["muted"], labelsize=9, length=0)

    def _pie_items(self, cd: pd.DataFrame) -> List[Tuple[str, float, str]]:
        val_col = [c for c in cd.columns if str(c).startswith("Значение, млн")][0]
        ramp_map = {"11": self.style["gray_ramp"],
                    "12": self.style["green_ramp"],
                    "13": self.style["green_ramp"]}
        recs = []
        for _, r in cd.iterrows():
            v = _num(r[val_col])
            if v is None or v <= 0:
                continue
            code = str(r["Код"])
            ramp = ramp_map.get(code[:2], self.style["gray_ramp"])
            recs.append({"Статья": str(r["Статья"]), "Код": code,
                         "Значение": v, "Цвет": ramp[len(recs) % len(ramp)]})
        if not recs:
            return []
        dups = {name for name, n in
                Counter(x["Статья"] for x in recs).items() if n > 1}
        for x in recs:
            if x["Статья"] in dups:
                x["Статья"] = f"{x['Статья']} (стр. {x['Код']})"
        df = pd.DataFrame(recs).sort_values("Значение", ascending=False)
        total = df["Значение"].sum()
        if df.empty or not total:
            return []
        return [(r["Статья"], float(r["Значение"]), r["Цвет"])
                for _, r in df.iterrows()]

    def _pie_png(self, name: str) -> Optional[str]:
        cd = self.results.get("chart_data", {}).get(name)
        if cd is None:
            return None
        items = self._pie_items(cd)
        if not items:
            return None
        # Объединение статей <1% в «Прочие <1%» + краска: топ-3 акцентные,
        # остальные — серые оттенки (решение 30.10.2026)
        kept = merge_small([(s, v) for s, v, _ in items])
        colors = accent_colors(kept, self.style)
        labels = [s for s, _ in kept]
        values = [v for _, v in kept]
        fig, ax = plt.subplots(figsize=(7.2, 4.8))
        wedges, _, autotexts = ax.pie(
            values, colors=colors, startangle=90, counterclock=False,
            autopct=lambda p: f"{p:.0f}%".replace(".", ","),
            pctdistance=0.72,
            wedgeprops={"edgecolor": "white", "linewidth": 2})
        for t in autotexts:
            t.set_fontsize(9)
        ax.text(0, 1.14, f"{name}, млн руб.", ha="center", fontsize=14,
                fontweight="bold", color=self.style["ink"])
        legend = [f"{l} — {ru_num(v)} млн руб." for l, v in
                  zip(labels, values)]
        ax.legend(wedges, legend, loc="center left", bbox_to_anchor=(1.0, 0.5),
                  fontsize=8, frameon=False, labelcolor=self.style["muted"])
        ax.axis("equal")
        path = os.path.join(self.out_dir, f"{name.replace(' ', '_')}.png")
        fig.savefig(path, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        self._imgs[name] = path
        return path

    def _dynamics_png(self) -> Optional[str]:
        cd = self.results.get("chart_data", {}).get(
            "Динамика изменений статей Баланса")
        if cd is None:
            return None
        rows = []
        for _, r in cd.iterrows():
            ch = _num(r["Изменение, млн руб."])
            if ch is None:
                continue
            pct = str(r.get("Изменение, %", ""))
            label = (f"{ch:+,.1f}".replace(",", " ") +
                     (f" ({pct}%)" if pct and pct != MISSING else ""))
            rows.append((str(r["Статья"]), ch, label))
        if not rows:
            return None
        rows = rows[::-1]
        fig, ax = plt.subplots(figsize=(7.8, max(2.4, 0.34 * len(rows) + 0.9)))
        ax.set_facecolor(self.style["surface"])
        names = [r[0] for r in rows]
        vals = [r[1] for r in rows]
        bars = ax.barh(names, vals, height=0.62, color=self.style["dynamics"],
                       zorder=2)
        vmax = max(abs(v) for v in vals) or 1.0
        for bar, (_, _, label) in zip(bars, rows):
            w = bar.get_width()
            if w >= 0:
                ax.text(w + vmax * 0.01, bar.get_y() + bar.get_height() / 2,
                        label, va="center", ha="left", fontsize=8,
                        color=self.style["muted"])
            else:
                ax.text(vmax * 0.01, bar.get_y() + bar.get_height() / 2,
                        label, va="center", ha="left", fontsize=8,
                        color=self.style["muted"])
        ax.set_xlim(min(0, min(vals) * 1.35), max(0, vmax * 1.4))
        ax.axvline(0, color=self.style["muted"], lw=0.8, zorder=1)
        ax.set_xticks([])
        self._strip_axes(ax)
        ax.set_title("Динамика изменений статей Баланса, млн руб.",
                     fontsize=14, fontweight="bold", loc="left",
                     color=self.style["ink"], pad=10)
        path = os.path.join(self.out_dir, "Динамика_статей.png")
        fig.savefig(path, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        self._imgs["Динамика изменений статей Баланса"] = path
        return path

    def _fin_results_png(self) -> Optional[str]:
        cd = self.results.get("chart_data", {}).get("Финансовые результаты")
        if cd is None:
            return None
        val_cols = [c for c in cd.columns if str(c).startswith("Значение, млн")]
        years = [str(c).rsplit("(", 1)[1].rstrip(")") for c in val_cols]
        cats = [str(r["Статья"]) for _, r in cd.iterrows()][::-1]
        data = {y: [_num(r[col]) for _, r in cd.iterrows()][::-1]
                for col, y in zip(val_cols, years)}
        if all(all(v is None for v in data[y]) for y in years):
            return None
        fig, ax = plt.subplots(figsize=(7.8, max(3.2, 0.42 * len(cats) + 1.2)))
        ax.set_facecolor(self.style["surface"])
        all_vals = [v for y in years for v in data[y] if v is not None]
        vmax = max((abs(v) for v in all_vals), default=1.0) or 1.0
        bar_h = 0.36 if len(years) > 1 else 0.6
        for i, y in enumerate(years):
            if len(years) > 1:
                pos = [j + (bar_h / 2 + 0.02) * (1 if i == 0 else -1)
                       for j in range(len(cats))]
            else:
                pos = list(range(len(cats)))
            for p, v in zip(pos, data[y]):
                if v is None:
                    continue
                ax.barh(p, v, height=bar_h,
                        color=self.style["series"][i % len(self.style["series"])],
                        zorder=2)
                x = v + vmax * 0.01 if v >= 0 else v - vmax * 0.01
                ax.text(x, p, ru_num(v), va="center",
                        ha="left" if v >= 0 else "right", fontsize=8,
                        color=self.style["muted"])
        ax.set_xlim(min(0, min(all_vals) * 1.35),
                    max(0, max(all_vals) * 1.4))
        ax.axvline(0, color=self.style["muted"], lw=0.8, zorder=1)
        ax.set_yticks(range(len(cats)))
        ax.set_yticklabels([c.replace(" (стр. ", "\n(стр. ") for c in cats],
                           fontsize=8)
        ax.set_xticks([])
        self._strip_axes(ax)
        ax.set_title("Финансовые результаты, млн руб.", fontsize=14,
                     fontweight="bold", loc="left", color=self.style["ink"],
                     pad=10)
        if len(years) > 1:
            from matplotlib.patches import Patch
            handles = [Patch(facecolor=self.style["series"][i], label=y)
                       for i, y in enumerate(years)]
            ax.legend(handles=handles, loc="lower right", fontsize=8,
                      frameon=False, title="Год",
                      labelcolor=self.style["muted"])
        path = os.path.join(self.out_dir, "Фин_результаты.png")
        fig.savefig(path, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        self._imgs["Финансовые результаты"] = path
        return path

    def _kpi_group_png(self, title: str, names_map: Dict[str, str],
                       nd: int) -> Optional[str]:
        by_name = {r["name"]: r for r in self.results.get("kpi", [])}
        years = sorted({y for r in self.results["kpi"] for y in r["values"]},
                       reverse=True)
        cats, data = [], {str(y): [] for y in years}
        for name, short in names_map.items():
            row = by_name.get(name)
            if row is None:
                continue
            if all(row["values"].get(y) is None for y in years):
                continue
            cats.append(short)
            for y in years:
                data[str(y)].append(row["values"].get(y))
        if not cats:
            return None
        fig, ax = plt.subplots(figsize=(7.6, 3.6))
        ax.set_facecolor(self.style["surface"])
        n = len(years)
        width = 0.72 / n
        all_vals = [v for row in data.values() for v in row if v is not None]
        lo, hi = (min(all_vals), max(all_vals)) if all_vals else (0.0, 1.0)
        span = (hi - lo) or max(abs(hi), 1.0)
        for i, y in enumerate(years):
            offs = [j - 0.36 + width / 2 + i * width for j in range(len(cats))]
            xs = [o for o, v in zip(offs, data[str(y)]) if v is not None]
            vs = [v for v in data[str(y)] if v is not None]
            ax.bar(xs, vs, width=width * 0.92,
                   color=self.style["series"][i % len(self.style["series"])],
                   label=str(y), zorder=2)
            for o, v in zip(xs, vs):
                ax.text(o, v + span * 0.03 * (1 if v >= 0 else -1),
                        ru_num(v, nd), ha="center",
                        va="bottom" if v >= 0 else "top",
                        fontsize=8, color=self.style["muted"])
        ax.axhline(0, color=self.style["muted"], lw=0.8, zorder=1)
        ax.set_xticks(range(len(cats)))
        ax.set_xticklabels(cats, fontsize=8, color=self.style["muted"])
        ax.set_yticks([])
        ax.set_ylim(min(0, lo - span * 0.25), hi + span * 0.3)
        self._strip_axes(ax)
        ax.set_title(title, fontsize=14, fontweight="bold", loc="left",
                     color=self.style["ink"], pad=10)
        ax.legend(loc="upper center", ncols=min(n, 3), fontsize=8,
                  frameon=False, labelcolor=self.style["muted"])
        key = title.replace(" ", "_").replace(",", "")
        path = os.path.join(self.out_dir, f"KPI_{key}.png")
        fig.savefig(path, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        self._imgs[title] = path
        return path

    def render_all(self) -> Dict[str, str]:
        for name in ("Структура активов", "Структура капитала"):
            self._pie_png(name)
        self._dynamics_png()
        self._fin_results_png()
        for title, (nd,) in KPI_KIND.items():
            self._kpi_group_png(title, KPI_GROUPS[title], nd)
        return dict(self._imgs)
class Content:
    """Содержательные тексты: выводы, влияние пассивов, рекомендации.

    Все расчетные числа берутся из результатов AnalyticsEngine;
    оценка «что хорошо / что плохо» — по правилам на динамику Баланса.
    """

    def __init__(self, results: Dict[str, Any]):
        self.results = results
        self.by_name = {r["name"]: r for r in results["kpi"]}
        self.years = sorted({y for r in results["kpi"] for y in r["values"]},
                            reverse=True)
        self._v = None

    def set_value_getter(self, fn):
        """fn(form, code, year) -> значение (тыс. руб.) для рекомендаций."""
        self._v = fn

    def _latest(self, name):
        row = self.by_name.get(name)
        if row is None:
            return None
        for y in sorted(row["values"].keys(), reverse=True):
            if row["values"].get(y) is not None:
                return y, row["values"][y]
        return None

    def _year_label(self) -> str:
        return str(self.years[0]) if self.years else "требует уточнения"

    def _prev_label(self) -> str:
        return f"{self.years[0]} к {self.years[1]} г." if len(self.years) > 1 \
            else ""

    def _dyn_rows(self) -> List[Tuple[str, float]]:
        """(статья, изменение) из Баланса (в млн, без промежуточных итогов)."""
        cd = self.results.get("chart_data", {}).get(
            "Динамика изменений статей Баланса")
        rows = []
        if cd is not None:
            for _, r in cd.iterrows():
                v = _num(r["Изменение, млн руб."])
                if v is not None:
                    rows.append((str(r["Статья"]), v))
        return rows

    def _pct_of(self, article: str) -> str:
        cd = self.results.get("chart_data", {}).get(
            "Динамика изменений статей Баланса")
        if cd is not None:
            sel = cd[cd["Статья"].astype(str) == article]
            if not sel.empty:
                p = str(sel.iloc[0].get("Изменение, %", ""))
                if p and p != MISSING:
                    return f" ({p}%)"
        return ""

    # -------- Выводы по 6 разделам методологии (docx, краткие) ------------ #
    def conclusions(self) -> List[str]:
        year = self._year_label()
        out = []
        names = ["Коэффициент текущей ликвидности",
                 "Коэффициент быстрой ликвидности",
                 "Коэффициент абсолютной ликвидности",
                 "Коэффициент обеспеченности собственными источниками"]
        parts = []
        for name in names:
            lv = self._latest(name)
            short = name.replace("Коэффициент ", "").lower()
            row = self.by_name.get(name)
            parts.append(f"коэффициент {short} — {ru_num(lv[1], 2)} "
                         f"({status_word(row['status'])})" if lv
                         else f"коэффициент {short} — {MISSING}")
        out.append(f"Ликвидность и платежеспособность за {year} г.: "
                   + "; ".join(parts) + ".")

        parts = []
        for name in ("Оборачиваемость дебиторской задолженности, дней",
                     "Оборачиваемость кредиторской задолженности, дней",
                     "Оборачиваемость запасов, дней",
                     "Оборачиваемость оборотного капитала, дней",
                     "Цикл конверсии денежных средств, дней"):
            lv = self._latest(name)
            short = name.replace("Оборачиваемость ", "").replace(", дней", "")
            parts.append(f"{short} — {ru_num(lv[1], 0)} дн." if lv
                         else f"{short} — {MISSING}")
        out.append(f"Деловая активность за {year} г.: " + "; ".join(parts)
                   + ".")

        parts = []
        for name in ("Рентабельность продаж (ROS), %",
                     "Рентабельность активов (ROA), %",
                     "Рентабельность собственного капитала (ROE), %"):
            lv = self._latest(name)
            short = name.split("(")[1].rstrip("), %")
            parts.append(f"{short} — {ru_num(lv[1], 1)} %" if lv
                         else f"{short} — {MISSING}")
        out.append(f"Рентабельность за {year} г.: " + "; ".join(parts) + ".")

        parts = []
        for name in ("Debt/Equity", "Коэффициент финансовой устойчивости"):
            lv = self._latest(name)
            row = self.by_name.get(name)
            parts.append(f"{name} — {ru_num(lv[1], 2)} "
                         f"({status_word(row['status'])})" if lv
                         else f"{name} — {MISSING}")
        out.append(f"Структура капитала и долговая нагрузка за {year} г.: "
                   + "; ".join(parts) + ".")

        out.append(self.horizontal_conclusion())

        parts = []
        z5 = self._latest("Z-счет Альтмана (пятифакторная модель)")
        row5 = self.by_name.get("Z-счет Альтмана (пятифакторная модель)")
        z2 = self._latest("Z-счет Альтмана (двухфакторная модель)")
        taf = self._latest("Модель Таффлера")
        rowt = self.by_name.get("Модель Таффлера")
        parts.append(f"Z-счет Альтмана (пятифакторная модель) = "
                     f"{ru_num(z5[1], 2)} — зона "
                     f"«{status_word(row5['status']).lower()}»" if z5
                     else f"Z-счет Альтмана (пятифакторная модель) — {MISSING}")
        parts.append(f"двухфакторная модель = {ru_num(z2[1], 2)} "
                     f"({'вероятность банкротства < 50%' if z2[1] < 0 else 'вероятность банкротства > 50%'})"
                     if z2 else f"двухфакторная модель — {MISSING}")
        parts.append(f"модель Таффлера = {ru_num(taf[1], 2)} — статус "
                     f"«{status_word(rowt['status']).lower()}»" if taf
                     else f"модель Таффлера — {MISSING}")
        out.append(f"Риски банкротства за {year} г.: " + "; ".join(parts)
                   + ".")
        return [anonymize(t) for t in out]

    def horizontal_conclusion(self) -> str:
        rows = [r for r in self._dyn_rows() if r[0] != "БАЛАНС (пассив)"]
        if not rows:
            return (f"Горизонтальный анализ: {MISSING} — данных достаточно "
                    f"только за один период.")
        rows.sort(key=lambda t: -abs(t[1]))
        txt = "; ".join(
            f"{name}: {v:+,.1f} млн".replace(",", " ") + self._pct_of(name)
            for name, v in rows[:3])
        return (f"Горизонтальный анализ ({self._prev_label()}): наиболее "
                f"существенные изменения статей Баланса: {txt}. Динамика "
                f"ОФР, ОДДС и ОИК — на листе «Горизонтальный анализ» файла "
                f"explonation.xlsx.")

    # ---- Влияние изменений пассивов на статьи актива (п. 3.7 docx) ------- #
    def passive_influence(self) -> str:
        cd = self.results.get("chart_data", {}).get(
            "Динамика изменений статей Баланса")
        if cd is None:
            return (f"Влияние изменения статей Пассива на статьи Актива: "
                    f"{MISSING} — данных достаточно только за один период.")
        items = [(str(r["Код"]), str(r["Статья"]),
                  _num(r["Изменение, млн руб."]))
                 for _, r in cd.iterrows()]
        act = [(n, v) for c, n, v in items if v is not None and c[:2] in
               ("11", "12") and c != "1700"]
        pas = [(n, v) for c, n, v in items if v is not None and c[:2] in
               ("13", "14", "15")]
        if not act or not pas:
            return (f"Влияние изменения статей Пассива на статьи Актива: "
                    f"{MISSING}.")

        def fmt(sel) -> str:
            sel = sorted(sel, key=lambda t: -abs(t[1]))[:3]
            return "; ".join(f"{n} {v:+,.1f}".replace(",", " ")
                             + self._pct_of(n) for n, v in sel) or "—"

        net_pas = sum(v for _, v in pas)
        net_act = sum(v for _, v in act)
        pas_str = fmt(pas)
        act_str = fmt(act)
        text = (f"Влияние изменения статей Пассива Баланса на статьи Актива "
                f"({self._prev_label()}): изменение источников — "
                f"{net_pas:+,.1f} млн".replace(",", " ") + f" ({pas_str}); "
                f"изменение активов — "
                f"{net_act:+,.1f} млн".replace(",", " ")
                + f" ({act_str}). Средства, привлеченные и накопленные в "
                  f"Пассиве, израсходованы на изменение статей Актива "
                  f"в полном объеме (Баланс сходится: "
                  f"{net_pas:+,.1f} = "
                  f"{net_act:+,.1f} млн).".replace(",", " "))
        good, bad = self.good_bad()
        if good:
            text += " ХОРОШО: " + "; ".join(good) + "."
        if bad:
            text += " ПЛОХО: " + "; ".join(bad) + "."
        return anonymize(text)

    def good_bad(self) -> Tuple[List[str], List[str]]:
        """Оценка «что хорошо / что плохо» по правилам на динамику Баланса."""
        rows = self._dyn_rows()
        if not rows:
            return [], []
        d = {n: v for n, v in rows}
        good, bad = [], []

        def ch(sub: str) -> Optional[Tuple[str, float]]:
            for n, v in rows:
                if sub in n:
                    return n, v
            return None

        kz = [t for t in rows if "Кредиторская задолженность" in t[0]
              and t[1] > 0]
        if kz:
            bad.append(
                f"рост активов профинансирован преимущественно кредиторской "
                f"задолженностью ({kz[0][1]:+,.1f} млн".replace(",", " ")
                + self._pct_of(kz[0][0]) + ") — бесплатный, но рискованный "
                "источник: при погашении платежи предстоит найти")
        dsz = ch("Дебиторская задолженность")
        if dsz and dsz[1] > 0:
            bad.append(
                f"дебиторская задолженность выросла ({dsz[1]:+,.1f} "
                "млн".replace(",", " ") + self._pct_of(dsz[0]) + ") — "
                "средства заморожены в расчетах с покупателями")
        sk = ch("Собственный капитал")
        if sk and sk[1] > 0:
            good.append(
                f"собственный капитал вырос ({sk[1]:+,.1f} млн".replace(
                    ",", " ") + self._pct_of(sk[0]) + ") — накопление "
                "устойчивых источников развития")
        cash = ch("Денежные средства")
        if cash and cash[1] > 0:
            good.append(
                f"денежные средства увеличились ({cash[1]:+,.1f} млн".replace(
                    ",", " ") + self._pct_of(cash[0]) + ") — запас "
                "ликвидности")
        fin = ch("Финансовые вложения (за исключением")
        if fin and fin[1] > 0:
            good.append(
                f"финансовые вложения выросли ({fin[1]:+,.1f} млн".replace(
                    ",", " ") + self._pct_of(fin[0]) + ") — размещение "
                "свободных средств")
        return good, bad

    def concrete_recommendations(self) -> List[str]:
        """Числовые целевые рекомендации (млн руб.), без коэффициентов.

        Каждая рекомендация — конкретное действие с суммой: до какого
        значения изменить статью, чтобы норматив достигнут.
        Требует value_getter(form, code, year) -> тыс. руб. (set_content_v).
        """
        recs: List[str] = []
        if self._v is None:
            return [MISSING + " — нет данных для расчета рекомендаций"]
        year = self._year_label()
        if year == MISSING:
            return recs
        y = int(year)

        def v(code):
            return self._v("Баланс", code, y)

        def fin(code):
            return self._v("ОФР", code, y)

        def m(x):
            return None if x is None else x / 1000

        oa, ko = v("1200"), v("1500")
        if oa is not None and ko is not None and ko > 0 and oa / ko < 1.5:
            target_ko = oa / 1.5
            delta = (ko - target_ko) / 1000
            if delta > 0:
                recs.append(f"Уменьшить кредиторскую задолженность на "
                            f"{ru_num(delta)} млн руб. (краткосрочные "
                            f"обязательства — до {ru_num(target_ko / 1000)} "
                            f"млн), чтобы довести текущую ликвидность до 1,5.")
        cash, cash_ko = v("1250"), v("1500")
        if cash is not None and cash_ko is not None and cash_ko > 0 \
                and cash / cash_ko < 0.3:
            delta = (0.3 * cash_ko - cash) / 1000
            if delta > 0:
                recs.append(f"Увеличить денежные средства на {ru_num(delta)} "
                            f"млн руб. (включая размещение свободных средств "
                            f"в ликвидные финансовые вложения), чтобы довести "
                            f"абсолютную ликвидность до 0,3.")
        sk, kba = v("1300"), v("1700")
        if sk is not None and kba is not None and kba > 0 and sk / kba < 0.5:
            delta = (0.5 * kba - sk) / 1000
            if delta > 0:
                recs.append(f"Увеличить собственный капитал на "
                            f"{ru_num(delta)} млн руб. (в первую очередь за "
                            f"счет накопления чистой прибыли), чтобы довести "
                            f"финансовую независимость до 0,5.")
        sk, do_, ko = v("1300"), v("1400"), v("1500")
        if sk is not None and do_ is not None and ko is not None and sk > 0 \
                and (do_ + ko) / sk > 1.0:
            delta = (do_ + ko - sk) / 1000
            if delta > 0:
                recs.append(f"Снизить заемные обязательства на "
                            f"{ru_num(delta)} млн руб. (до соотношения долга "
                            f"и капитала 1,0), начав с самого дорогого долга.")
        dz = v("1230")
        rev = fin("2110")
        if dz is not None and rev is not None and rev > 0:
            target_dz = rev * 100 / 365
            delta = (dz - target_dz) / 1000
            if delta > 0:
                recs.append(f"Сократить дебиторскую задолженность на "
                            f"{ru_num(delta)} млн руб. (до "
                            f"{ru_num(target_dz / 1000)} млн), чтобы довести "
                            f"ее оборачиваемость до 100 дней — за счет работы "
                            f"с просроченными расчетами и авансирования "
                            f"договоров.")
        ros_val = self._latest("Рентабельность продаж (ROS), %")
        sells, revv = fin("2200"), fin("2110")
        if ros_val is not None and ros_val[1] < 0 and sells is not None \
                and revv is not None:
            delta = (0.03 * revv - sells) / 1000
            recs.append(f"Повысить прибыль от продаж на {ru_num(delta)} млн "
                        f"руб. (до маржинальности 3% — {ru_num(0.03 * revv / 1000)} "
                        f"млн при текущей выручке) — за счет ценообразования "
                        f"и снижения себестоимости.")
        if not recs:
            recs.append("Существующих отклонений показателей от нормативов "
                        "не выявлено; поддерживать текущую финансовую "
                        "политику.")
        recs.append("Показатели, помеченные «требует уточнения», повторно "
                    "оценить после восполнения данных.")
        return [anonymize(t) for t in recs]
    # ---- Качественные выводы (без коэффициентов, финальный слайд pptx) ---- #
    def qualitative_conclusions(self) -> List[str]:
        out = []
        cr = self._latest("Коэффициент текущей ликвидности")
        if cr is not None and cr[1] < 1.5:
            out.append("Краткосрочная платежеспособность ограничена: "
                       "оборотных активов не хватает для комфортного "
                       "покрытия текущих долгов.")
        ros = self._latest("Рентабельность продаж (ROS), %")
        if ros is not None and ros[1] <= 0:
            out.append("Продажи не приносят прибыли: себестоимость "
                       "поглощает выручку, каждый новый контракт работает "
                       "в минус.")
        dz = self._latest("Оборачиваемость дебиторской задолженности, дней")
        kz = self._latest("Оборачиваемость кредиторской задолженности, дней")
        if dz is not None and kz is not None and dz[1] < kz[1]:
            out.append("Деньги застревают в дебиторской задолженности "
                       "дольше, чем поступают от поставщиков в кредит — "
                       "дефицит операционных денег растет.")
        de = self._latest("Debt/Equity")
        if de is not None and de[1] > 1.0:
            out.append("Долг превышает собственный капитал: устойчивость "
                       "зависит от доступа банков и кредиторов к "
                       "рефинансированию.")
        z5 = self._latest("Z-счет Альтмана (пятифакторная модель)")
        row5 = self.by_name.get("Z-счет Альтмана (пятифакторная модель)")
        if z5 is not None and row5 is not None:
            zone = status_word(row5["status"]).lower()
            if zone in ("риск", "красная"):
                out.append("Модели прогнозирования банкротства указывают "
                           "на зону риска — необходим план восстановления "
                           "финансовой устойчивости.")
        if not out:
            out.append("Финансовое состояние устойчиво: нормативы "
                       "соблюдены, значимых нарушений не выявлено.")
        return [anonymize(t) for t in out]

    # ---- Развёрнутые выводы по слайдам (2–4 предложения, bold) ------------ #
    def slide_conclusions(self) -> Dict[str, str]:
        """{заголовок диаграммы/раздела: текст 2–4 предложений}."""
        out: Dict[str, str] = {}

        def s(name, fmt="коэффициент", nd=2):
            lv = self._latest(name)
            row = self.by_name.get(name)
            if lv is None:
                return f"{name.replace('Коэффициент ', '')} — {MISSING}"
            status = status_word(row["status"]) if row else "—"
            if fmt == "коэффициент":
                return (f"{name.replace('Коэффициент ', '')} — "
                        f"{ru_num(lv[1], nd)} ({status})")
            return None

        # Ликвидность
        parts = []
        cur = self._latest("Коэффициент текущей ликвидности")
        row = self.by_name.get("Коэффициент текущей ликвидности")
        if cur is not None:
            st = status_word(row["status"])
            parts.append(
                f"Текущую ликвидность {ru_num(cur[1], 2)} оцениваем как "
                f"«{st}»: оборотные активы покрывают краткосрочные "
                f"обязательства на {ru_num(cur[1] * 100, 0)}%.")
        quick = self._latest("Коэффициент быстрой ликвидности")
        if quick is not None:
            parts.append(f"Без запасов покрытие рассчитывается на "
                         f"{ru_num(quick[1], 2)} при нормативе 0,8–1,0.")
        own = self._latest("Коэффициент обеспеченности собственными "
                           "источниками")
        if own is not None:
            parts.append(f"Собственными источниками покрыто "
                         f"{ru_num(own[1] * 100, 0)}% оборотных активов.")
        if cur is not None and cur[1] < 1.5:
            parts.append("Источники риска — быстрый рост краткосрочных "
                         "обязательств: платежеспособность ограничена.")
        out["Ликвидность"] = " ".join(parts[:4])

        # Финансовая устойчивость
        parts = []
        de = self._latest("Debt/Equity")
        row = self.by_name.get("Debt/Equity")
        if de is not None:
            st = status_word(row["status"])
            if de[1] > 1.0:
                parts.append(
                    f"Долг превышает собственный капитал: кредиторские и "
                    f"заемные средства в {ru_num(de[1], 1)} раза больше "
                    f"собственных источников.")
            else:
                parts.append(
                    f"Соотношение долга и капитала {ru_num(de[1], 2)} — "
                    f"«{st}»: устойчивость обеспечена.")
        stab = self._latest("Коэффициент финансовой устойчивости")
        if stab is not None:
            parts.append(f"Долгосрочные источники финансируют "
                         f"{ru_num(stab[1] * 100, 0)}% активов.")
        out["Финансовая устойчивость"] = " ".join(parts[:3])

        # Рентабельность
        parts = []
        ros = self._latest("Рентабельность продаж (ROS), %")
        if ros is not None:
            if ros[1] <= 0:
                parts.append("Продажи не приносят прибыли: себестоимость "
                             "поглощает выручку (ROS отрицателен).")
            else:
                parts.append(f"С каждого рубля продаж компания "
                             f"зарабатывает {ru_num(ros[1], 1)} коп. "
                             f"операционной прибыли.")
        roa = self._latest("Рентабельность активов (ROA), %")
        if roa is not None:
            parts.append(f"Активы генерируют {ru_num(roa[1], 1)}% чистой "
                         f"прибыли за год — эффективность низкая.")
        roe = self._latest("Рентабельность собственного капитала (ROE), %")
        if roe is not None:
            parts.append(f"Доходность средств собственников — "
                         f"{ru_num(roe[1], 1)}%.")
        out["Рентабельность"] = " ".join(parts[:4])

        # Оборачиваемость
        parts = []
        dz = self._latest("Оборачиваемость дебиторской задолженности, дней")
        kz = self._latest("Оборачиваемость кредиторской задолженности, дней")
        if dz is not None:
            parts.append(f"Покупатели рассчитываются за "
                         f"{ru_num(dz[1], 0)} дней.")
        if kz is not None:
            parts.append(f"Сроки расчетов с поставщиками — "
                         f"{ru_num(kz[1], 0)} дней.")
        if dz is not None and kz is not None and dz[1] < kz[1]:
            parts.append("Разрыв в "
                         f"{ru_num(kz[1] - dz[1], 0)} дней покрывает "
                         "операционный цикл за счет кредиторов.")
        ccc = self._latest("Цикл конверсии денежных средств, дней")
        if ccc is not None:
            parts.append(f"Полный цикл оборота денег составляет "
                         f"{ru_num(ccc[1], 0)} дней.")
        out["Оборачиваемость, дней"] = " ".join(parts[:4])

        # Структуры
        cd = self.results.get("chart_data", {}).get("Структура активов")
        if cd is not None:
            val_col = [c for c in cd.columns
                       if str(c).startswith("Значение, млн")][0]
            total = sum(_num(r[val_col]) for _, r in cd.iterrows()
                        if _num(r[val_col]) is not None)
            parts = []
            if total:
                top = cd.sort_values(val_col, ascending=False).iloc[0]
                parts.append(
                    f"Активы сконцентрированы в статье "
                    f"«{top['Статья']}» — {ru_num(_num(top[val_col]) / total * 100, 0)}% "
                    f"от итога баланса.")
                out["Структура активов"] = " ".join(parts[:2])
        cd = self.results.get("chart_data", {}).get("Структура капитала")
        if cd is not None:
            val_col = [c for c in cd.columns
                       if str(c).startswith("Значение, млн")][0]
            total = sum(_num(r[val_col]) for _, r in cd.iterrows()
                        if _num(r[val_col]) is not None)
            kz_v = sum(_num(r[val_col]) for _, r in cd.iterrows()
                       if str(r["Код"]) == "1520") or 0
            parts = []
            if total:
                parts.append(
                    f"Кредиторская задолженность составляет "
                    f"{ru_num(kz_v / total * 100, 0)}% источников — "
                    "структура капитала перегружена краткосрочными "
                    "долгами.")
                parts.append("Рекомендуется перегруппировка источников: "
                             "часть коротких долгов — в долгосрочные "
                             "инструменты.")
            out["Структура капитала"] = " ".join(parts[:4])

        # Динамика статей Баланса
        rows = [r for r in self._dyn_rows() if r[0] != "БАЛАНС (пассив)"]
        if rows:
            rows.sort(key=lambda t: -abs(t[1]))
            name, v = rows[0]
            out["Динамика изменений статей Баланса"] = (
                f"Наибольшее изменение — статья «{name}»: "
                f"{v:+,.1f} млн".replace(",", " ") + self._pct_of(name) + ". "
                "Изменение активов профинансировано за счет источников "
                "(см. слайд «Структура капитала»): структура баланса "
                "приведена в соответствие с новым масштабом бизнеса.")

        # Финансовые результаты
        fin = self.results.get("chart_data", {}).get(
            "Финансовые результаты")
        if fin is not None:
            val_cols = [c for c in fin.columns
                        if str(c).startswith("Значение, млн")]
            if len(val_cols) >= 2:
                cur_c, prev_c = val_cols[0], val_cols[1]
                rows_f = {str(r["Статья"]): (_num(r[cur_c]), _num(r[prev_c]))
                          for _, r in fin.iterrows()}
                parts = []
                for label in ("Выручка",):
                    a, b = rows_f.get(label, (None, None))
                    if a is not None and b is not None and b:
                        pct = f"{(a - b) / abs(b) * 100:+.0f}%"
                        parts.append(
                            f"Выручка изменилась с {ru_num(b)} на "
                            f"{ru_num(a)} млн ({pct}).")
                for label in ("Чистая прибыль (убыток)",):
                    a, b = rows_f.get(label, (None, None))
                    if a is not None and b is not None:
                        parts.append(
                            f"Чистая прибыль изменилась с {ru_num(b)} на "
                            f"{ru_num(a)} млн: маржинальность снизилась, "
                            "запас прочности ограничен.")
                out["Финансовые результаты"] = " ".join(parts[:4])

        # Риски банкротства
        parts = []
        z5 = self._latest("Z-счет Альтмана (пятифакторная модель)")
        row5 = self.by_name.get("Z-счет Альтмана (пятифакторная модель)")
        if z5 is not None and row5 is not None:
            parts.append(
                f"Z-счет Альтмана {ru_num(z5[1], 2)} — зона "
                f"«{status_word(row5['status']).lower()}»: вероятность "
                f"банкротства по пятифакторной модели оценивается "
                f"системно.")
        taf = self._latest("Модель Таффлера")
        rowt = self.by_name.get("Модель Таффлера")
        if taf is not None and rowt is not None:
            parts.append(
                f"Модель Таффлера: {ru_num(taf[1], 2)} — "
                f"«{status_word(rowt['status']).lower()}».")
        z2 = self._latest("Z-счет Альтмана (двухфакторная модель)")
        if z2 is not None:
            parts.append(
                "Двухфакторная модель: вероятность банкротства "
                f"{'< 50% (умеренная)' if z2[1] < 0 else '> 50% (высокая)'}.")
        out["Риски банкротства"] = " ".join(parts[:4])
        return {k: anonymize(v) for k, v in out.items()}

SRC_SHEET = "Исходные данные"
CALC_SHEET = "Расчеты"
HORIZ_SHEET = "Горизонтальный анализ"
DATA_SHEET = "Данные для диаграмм"
CHARTS_SHEET = "Диаграммы"

_PL_RE = re.compile(r"\{([A-Z]+)(?:[@:]([^}]*))?\}")


class ReportGenerator:
    """Генерация explonation.xlsx, rezult.docx, presentation.pptx (AC-04/05/09).

    Все расчетные значения на листах «Расчеты» и «Горизонтальный анализ»
    записываются Excel-формулами со ссылками на «Исходные данные» (и
    промежуточные строки «Расчетов»), каждое значение — в отдельной ячейке;
    диаграммы строятся на области «Данные для диаграмм» и остаются
    нативными (редактируемыми) объектами Excel. В Word диаграммы
    скопированы из того же файла как OLE-объекты (через COM; при
    недоступном Excel — PNG). При создании файла сразу выполняется
    контрольная проверка анонимизации."""

    def __init__(self, results: Dict[str, Any], forms: Dict[str, List[dict]],
                 style: Optional[Dict[str, Any]] = None,
                 dep: Optional[Dict[int, float]] = None):
        self.results = results
        self.forms = forms
        self.style = style or {}
        self.dep = dep or {}
        self._charts: Dict[str, str] = {}
        self._charts_dir: Optional[str] = None
        # цвета секторов круговых: [имя pie-блока] -> [hex по строкам блока]
        self._pie_colors: Dict[str, List[str]] = {}
        engine = AnalyticsEngine(forms)
        self._v = engine.v                      # сырые значения (тыс. руб.)
        self.years = engine.years
        self.cell_map: Dict[tuple, str] = {}    # (form, code, colkey) -> $Col$Row
        self.kcells: Dict[tuple, str] = {}      # (kpi_name, str(year)) -> $Col$Row
        self.content = Content(results)

    # ------------------------------------------------------------------ #
    def _years_desc(self) -> List[int]:
        return sorted({y for r in self.results["kpi"] for y in r["values"]},
                      reverse=True)

    def _year_label(self) -> str:
        years = self._years_desc()
        return str(years[0]) if years else "требует уточнения"

    def _ensure_charts(self) -> Dict[str, str]:
        """PNG fallback для docx (если Excel COM недоступен)."""
        if self._charts:
            return self._charts
        self._charts_dir = self._charts_dir or tempfile.mkdtemp(
            prefix="fsa_charts_")
        renderer = ChartRenderer(self.results, self.style, self._charts_dir)
        self._charts = renderer.render_all()
        return self._charts

    # --- адреса исходных ячеек ------------------------------------------- #
    def _src(self, form: str, code: str, year: int,
             colkey: Optional[str] = None) -> Optional[str]:
        """Адрес ячейки «Исходных данных» или None, если строки/значения нет."""
        if colkey is not None:
            # ОИК (столбец «Итого» и т.п.)
            return self.cell_map.get((form, code, colkey))
        return self.cell_map.get((form, code, f"val_{year}"))

    def _resolve(self, tpl: str, year: int) -> Optional[str]:
        """Подстановка адресов в шаблон формулы ({B@1200}, {AVG@Баланс@1230},
        {DEP}, {K@имя}); None — если данных нет (в ячейку пишется
        «требует уточнения»)."""
        missing = [False]

        def ref(form: str, code: str, y: int) -> str:
            a = self._src(form, code, y)
            if a is None:
                missing[0] = True
                return ""
            return f"'{SRC_SHEET}'!{a}"

        def repl(m):
            kind, rest = m.group(1), (m.group(2) or "").strip()
            if kind == "DEP":
                a = self._src("Введенные данные", "АМОРТ", year)
                if a is None:
                    missing[0] = True
                    return ""
                return f"'{SRC_SHEET}'!{a}"
            if kind == "K":
                cell = self.kcells.get((rest, str(year)))
                if cell is None:
                    missing[0] = True
                    return ""
                return f"{CALC_SHEET}!{cell}"
            if kind == "AVG":
                try:
                    form, code = rest.split("@")
                except ValueError:
                    missing[0] = True
                    return ""
                cur = ref(form, code, year)
                if missing[0]:
                    return ""
                prev_a = self.cell_map.get((form, code, f"val_{year - 1}"))
                if prev_a is None:
                    return cur          # нет прошлого года — значение на дату
                # скобки обязательны: подстановка входит в знаменатель
                return f"(({cur}+'{SRC_SHEET}'!{prev_a})/2)"
            form = {"B": "Баланс", "OF": "ОФР", "OD": "ОДДС"}[kind]
            return ref(form, rest, year)

        out = _PL_RE.sub(repl, tpl)
        if missing[0]:
            return None
        return out

    # =====================================================================
    # Excel (FR-06, AC-04, AC-06) — все значения формулами
    # =====================================================================
    def generate_excel(self, path: str) -> None:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        bold = Font(bold=True)
        fill = PatternFill("solid", fgColor="EDEDED")
        years_all = self._years_desc()

        # --- Лист «Исходные данные» (+ карта адресов) ------------------- #
        ws = wb.active
        ws.title = SRC_SHEET
        row = 0
        for form, rows in self.forms.items():
            if not rows:
                continue
            if form == "ОИК":
                from data_extraction import OIK_CAPITAL_COLS
                present = {c for r in rows for c in r
                           if c not in ("indicator", "code", "section")}
                colkeys = [c for c in OIK_CAPITAL_COLS if c in present]
                colkeys += [c for c in present if c not in colkeys]
                col_names = list(colkeys)
            else:
                colkeys = [f"val_{y}" for y in years_all]
                col_names = [f"{y} г." for y in years_all]
            row += 1
            ws.append(["Форма", "Раздел", "Код", "Показатель", *col_names])
            for c in ws[row]:
                c.font = bold
                c.fill = fill
            for r in rows:
                row += 1
                code = str(r.get("code") or "")
                ws.append([form, anonymize(r.get("section")),
                           anonymize(code),
                           anonymize(str(r.get("indicator") or "")),
                           *[_num(r.get(ck)) for ck in colkeys]])
                if code:
                    for j, ck in enumerate(colkeys):
                        key = (form, code, ck)
                        if key not in self.cell_map:
                            # как в analytics.v — берется первая строка с кодом
                            self.cell_map[key] = \
                                f"${get_column_letter(5 + j)}${row}"

        # Введенные данные (амортизация) — для формул EBITDA
        if self.dep:
            colkeys = [f"val_{y}" for y in years_all]
            row += 1
            ws.append(["Введенные данные", "", "АМОРТ",
                       "Амортизация за период (введено аналитиком), тыс. руб.",
                       *[self.dep.get(y) if ck == f"val_{y}" else None
                         for y, ck in zip(years_all, colkeys)]])
            for c in ws[row]:
                if c.value is not None:
                    c.font = bold
            for i, ck in enumerate(colkeys):
                self.cell_map[("Введенные данные", "АМОРТ", ck)] = \
                    f"${get_column_letter(5 + i)}${row}"

        for col, w in (("A", 16), ("B", 30), ("C", 10), ("D", 50)):
            ws.column_dimensions[col].width = w
        for j in range(5, 5 + 8):
            ws.column_dimensions[get_column_letter(j)].width = 16

        # --- Лист «Расчеты»: значения формулами -------------------------- #
        ws2 = wb.create_sheet(CALC_SHEET)
        ws2.append(["Раздел", "Показатель", "Формула",
                    *[f"{y} г." for y in years_all],
                    "Диапазон (рекоменд.)", "Статус"])
        for c in ws2[1]:
            c.font = bold
            c.fill = fill

        # Блок «Основные показатели ФХД», млн руб. = src/1000
        for form, code, name in KEY_LINES:
            cells = []
            for y in years_all:
                a = self._src(form, code, y)
                cells.append(f"='{SRC_SHEET}'!{a}/1000" if a else MISSING)
            ws2.append(["Основные показатели ФХД", name,
                        f"({form}, стр. {code}) / 1000", *cells, "—", "—"])

        # Блок KPI: формулы из шаблонов analytics (placeholders -> адреса)
        for r in self.results["kpi"]:
            rownum = ws2.max_row + 1
            year_cells = []
            for j, y in enumerate(years_all):
                col_letter = get_column_letter(4 + j)
                expr = self._resolve(r["excel"], y) if r.get("excel") else None
                if expr is None:
                    year_cells.append(MISSING)
                else:
                    year_cells.append(f"=IFERROR(({expr}),\"\")")
                    self.kcells[(r["name"], str(y))] = \
                        f"${col_letter}${rownum}"
            # Статус формулой (BR-01): по последнему доступному году
            zone_fn_val = r.get("zone_fn")
            zone = "z5" if zone_fn_val is zzone_of else \
                ("taffler" if zone_fn_val is taffler_zone else None)
            norm = r.get("norm")
            status_formula = None
            if zone == "z5":
                v_expr = self._latest_year_expr(rownum, len(years_all))
                status_formula = (f"=IF(ISTEXT({v_expr}),"
                                  f"\"{MISSING}\",IF({v_expr}>2.9,"
                                  f"\"Безопасно\",IF({v_expr}>1.8,"
                                  f"\"Серая зона\",\"Риск\")))")
            elif zone == "taffler":
                v_expr = self._latest_year_expr(rownum, len(years_all))
                status_formula = (f"=IF(ISTEXT({v_expr}),"
                                  f"\"{MISSING}\",IF({v_expr}>0.3,"
                                  f"\"Безопасно\",IF({v_expr}>0.2,"
                                  f"\"Серая зона\",\"Риск\")))")
            elif norm is not None:
                lo, hi = norm
                v_expr = self._latest_year_expr(rownum, len(years_all))
                cond = ""
                if lo is not None and hi is not None:
                    cond = f"{v_expr}>={lo},{v_expr}<={hi}"
                elif lo is not None:
                    cond = f"{v_expr}>={lo}"
                elif hi is not None:
                    cond = f"{v_expr}<={hi}"
                status_formula = (f"=IF(ISTEXT({v_expr}),\"{MISSING}\","
                                  f"IF(AND({cond}),\"Безопасно\",\"Риск\"))")
            ws2.append([r["section"], r["name"], r["formula"], *year_cells,
                        r["range"],
                        status_formula if status_formula else "—"])

        ws2.column_dimensions["A"].width = 36
        ws2.column_dimensions["B"].width = 48
        ws2.column_dimensions["C"].width = 56
        for j in range(4, 4 + len(years_all)):
            ws2.column_dimensions[get_column_letter(j)].width = 14
        ws2.column_dimensions[
            get_column_letter(4 + len(years_all))].width = 26
        ws2.column_dimensions[get_column_letter(5 + len(years_all))].width = 16

        # Формат чисел для колонок лет
        for i in range(2, ws2.max_row + 1):
            nm = str(ws2.cell(row=i, column=2).value or "")
            fmt = "0"
            if ", дней" in nm:
                fmt = "0"
            elif "%" in nm:
                fmt = "0.0"
            elif "EBITDA" in nm or "Working capital" in nm \
                    or "млн руб." in str(ws2.cell(row=i, column=3).value):
                fmt = "0"
            else:
                fmt = "0.00"
            for j in range(4, 4 + len(years_all)):
                ws2.cell(row=i, column=j).number_format = fmt

        # --- Лист «Горизонтальный анализ» (формулы) ----------------------- #
        ws3 = wb.create_sheet(HORIZ_SHEET)
        pairs = [(years_all[i], years_all[i + 1])
                 for i in range(len(years_all) - 1)]
        from analytics import BALANCE_ASSET_TOTALS
        for form in ("Баланс", "ОФР", "ОДДС"):
            rows = [r for r in self.forms.get(form, [])
                    if (c := str(r.get("code") or ""))]
            if not rows:
                continue
            ws3.append([f"Форма: {form}"])
            ws3.cell(row=ws3.max_row, column=1).font = bold
            head = ["Код", "Показатель"]
            for cur, prev in pairs:
                head += [f"{cur} к {prev}, млн", f"{cur} к {prev}, %"]
            ws3.append(head)
            for c in ws3[ws3.max_row]:
                c.font = bold
                c.fill = fill
            for r in rows:
                code = str(r.get("code") or "")
                if form == "Баланс" and code in BALANCE_ASSET_TOTALS:
                    continue
                cells = [code, anonymize(str(r.get("indicator") or ""))]
                for cur, prev in pairs:
                    a = self._src(form, code, cur)
                    b = self._src(form, code, prev)
                    if a is None or b is None:
                        cells += [MISSING, MISSING]
                    else:
                        aa, bb = f"'{SRC_SHEET}'!{a}", f"'{SRC_SHEET}'!{b}"
                        cells += [f"=({aa}-{bb})/1000",
                                  f"=IF({bb}=0,\"—\",({aa}-{bb})/ABS({bb})"
                                  f"*100)"]
                ws3.append(cells)
            ws3.cell(row=ws3.max_row + 1, column=1)
        self._oik_horizontal(ws3)
        ws3.column_dimensions["A"].width = 10
        ws3.column_dimensions["B"].width = 44
        for j in range(3, 3 + len(pairs) * 2):
            ws3.column_dimensions[get_column_letter(j)].width = 18

        # --- Лист «Данные для диаграмм» (значения формулами) -------------- #
        block_ranges = self._fill_chart_data_sheet(wb)

        # --- Лист «Диаграммы» (нативные, источник — «Данные для диаграмм») #
        self._add_excel_charts(wb, block_ranges)

        wb.save(path)

    def _latest_year_expr(self, rownum: int, n_years: int) -> str:
        """Выражение Excel: первая (по порядку лет) ячейка с числом."""
        from openpyxl.utils import get_column_letter
        cols = [get_column_letter(4 + j) + str(rownum)
                for j in range(n_years)]
        expr = f"\"{MISSING}\""
        for col in reversed(cols):
            expr = f"IF(COUNT({col})>0,{col},{expr})"
        return expr

    def _oik_horizontal(self, ws3) -> None:
        """Блок ОИК: динамика капитала на отчетные даты, формулами."""
        from openpyxl.styles import Font, PatternFill
        bold = Font(bold=True)
        fill = PatternFill("solid", fgColor="EDEDED")
        dates = {2023: ("33100", "3100"), 2024: ("63200", "3200"),
                 2025: ("83300", "3300")}
        recs = []
        for y in self.years:
            closing, _ = dates.get(y, ("", ""))
            if closing:
                a = self.cell_map.get(("ОИК", closing, "Итого"))
                if a:
                    recs.append((y, a))
        if len(recs) < 2:
            return
        ws3.append(["Форма: ОИК (итого капитала, млн руб.)"])
        ws3.cell(row=ws3.max_row, column=1).font = bold
        ws3.append(["Период", "Капитал на конец, млн", "Изменение, млн",
                    "Изменение, %"])
        for c in ws3[ws3.max_row]:
            c.font = bold
            c.fill = fill
        for (y0, a0), (y1, a1) in zip(recs, recs[1:]):
            aa = f"'{SRC_SHEET}'!{a0}/1000"
            bb = f"'{SRC_SHEET}'!{a1}/1000"
            ws3.append([f"{y0} к {y1} (стр. {dates.get(y0)[0]}, "
                        f"{dates.get(y1)[0]})",
                        f"={aa}", f"={aa}-{bb}",
                        f"=IF({bb}=0,\"—\",({aa}-{bb})/{bb}*100)"])

    def _share_block(self, entries):
        """Строки сводного блока круговой: статьи с долей >=1% (в порядке
        исходных данных) + итоговая строка «Прочие <1%» (SUM).
        entries: (статья, адрес ячейки «Исходных данных», значение тыс. руб.).
        Возвращает (строки блока с формулами, пары (статья, значение))."""
        total = sum(v for _, _, v in entries)
        rows, small_addrs, small_sum = [], [], 0.0
        items = []
        for label, addr, v in entries:
            if total and v / total < 0.01:
                small_addrs.append(addr)
                small_sum += v
            else:
                rows.append((label, f"='{SRC_SHEET}'!{addr}/1000"))
                items.append((label, v))
        if small_addrs:
            refs = ",".join(
                f"'{SRC_SHEET}'!{a}" for a in small_addrs)
            rows.append((SMALL_LABEL, f"=SUM({refs})/1000"))
            items.append((SMALL_LABEL, small_sum))
        return rows, items

    # --- Данные для диаграмм ------------------------------------------------ #
    def _fill_chart_data_sheet(self, wb):
        """Блоки значений формулами; возвращает координаты блоков."""
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter
        ws = wb.create_sheet(DATA_SHEET)
        ws.cell(row=1, column=1,
                value="Данные для диаграмм: каждое значение — формула "
                      "со ссылкой на исходные данные/расчеты (можно "
                      "проверить). Диаграммы на листе «Диаграммы» построены "
                      "на этих областях.")
        ws.cell(row=1, column=1).font = Font(bold=True)

        ranges: Dict[str, Tuple[int, int, int, int]] = {}
        year0 = self._years_desc()[0]
        prev_y = self._years_desc()[1] if len(self._years_desc()) > 1 else None
        row = 3

        def put_block(title, headers, data_rows, cell_fmt=None):
            nonlocal row
            row += 1
            ws.cell(row=row, column=1, value=title).font = Font(bold=True)
            row += 1
            for j, h in enumerate(headers):
                ws.cell(row=row, column=1 + j, value=h)
            hrow = row
            first = row + 1
            for vals in data_rows:
                row += 1
                for j, v in enumerate(vals):
                    cell = ws.cell(row=row, column=1 + j, value=v)
                    # Формат ячеек значений: на него ориентируются метки
                    # данных диаграмм (dLbls numFmt Excel игнорирует)
                    if cell_fmt and j > 0 and v != MISSING:
                        cell.number_format = cell_fmt
            last = row
            row += 1
            return (1, hrow, 1 + len(headers) - 1, last), first, last

        # Структура активов/капитала (только строки с кодами в диапазоне)
        for name, lo, hi in (("Структура активов", 1110, 1269),
                             ("Структура капитала", 1310, 1549)):
            from analytics import BALANCE_ASSET_TOTALS, BALANCE_LIAB_TOTALS
            skip = BALANCE_ASSET_TOTALS if lo == 1110 else BALANCE_LIAB_TOTALS
            rows, entries = [], []
            for r in self.forms.get("Баланс", []):
                code = str(r.get("code") or "")
                if not code or code in skip:
                    continue
                if not lo <= int(code) <= hi:
                    continue
                a = self._src("Баланс", code, year0)
                if a is None:
                    continue
                label = anonymize(str(r.get("indicator") or ""))
                rows.append((label, f"='{SRC_SHEET}'!{a}/1000"))
                v = self._v("Баланс", code, year0)
                if v is not None and v > 0:
                    entries.append((label, a, v))
            if rows:
                ranges[name], _, _ = put_block(
                    f"{name}, млн руб. ({year0})", ["Статья", "Значение"],
                    rows, cell_fmt="#,##0.0")
                # Круговая рисуется по сводному блоку: статьи с долей >=1%,
                # остальное одной строкой «Прочие <1%» (решение 30.10.2026)
                pie_rows, pie_items = self._share_block(entries)
                if len(pie_rows) > 1:
                    pie_key = f"{name} (круговая)"
                    ranges[pie_key], _, _ = put_block(
                        f"{pie_key}, млн руб. ({year0}): статьи >=1% "
                        f"+ «{SMALL_LABEL}»", ["Статья", "Значение"],
                        pie_rows, cell_fmt="#,##0.0")
                    # Цвета секторов: топ-3 акцентные, остальные серые
                    self._pie_colors[pie_key] = accent_colors(
                        pie_items, self.style)

        # Динамика изменений статей Баланса
        if prev_y:
            rows = []
            for r in self.forms.get("Баланс", []):
                code = str(r.get("code") or "")
                name_txt = str(r.get("indicator") or "").strip().lower()
                if not code or code == "1600" \
                        or name_txt.startswith("итого по разделу"):
                    continue
                a = self._src("Баланс", code, year0)
                b = self._src("Баланс", code, prev_y)
                if a is None or b is None:
                    continue
                rows.append((anonymize(str(r.get("indicator") or "")),
                             f"=('{SRC_SHEET}'!{a}-'{SRC_SHEET}'!{b})/1000"))
            if rows:
                ranges["Динамика изменений статей Баланса"] = put_block(
                    "Динамика изменений статей Баланса, млн руб. "
                    f"({year0} к {prev_y})",
                    ["Статья", "Изменение"], rows,
                    cell_fmt="#,##0.0")[0]

        # Финансовые результаты (2 года)
        fin_rows = []
        for r in self.forms.get("ОФР", []):
            code = str(r.get("code") or "")
            if not code or int(code) > 2400:
                continue
            a = self._src("ОФР", code, year0)
            b = self._src("ОФР", code, prev_y) if prev_y else None
            if a is None or (prev_y and b is None):
                continue
            cells = [anonymize(str(r.get("indicator") or "")),
                     f"='{SRC_SHEET}'!{a}/1000"]
            if prev_y:
                cells.append(f"='{SRC_SHEET}'!{b}/1000")
            fin_rows.append(tuple(cells))
        if fin_rows:
            headers = ["Статья", f"{year0}"] + (
                [f"{prev_y}"] if prev_y else [])
            ranges["Финансовые результаты"], _, _ = put_block(
                "Финансовые результаты, млн руб.", headers, fin_rows,
                cell_fmt="#,##0.0")

        # KPI-группы (по годам, ссылки на «Расчеты»)
        for title, names_map in KPI_GROUPS.items():
            rows = []
            for name in names_map:
                cells = [names_map[name]]
                ok = False
                for y in self._years_desc():
                    cell = self.kcells.get((name, str(y)))
                    if cell is not None:
                        cells.append(f"={CALC_SHEET}!{cell}")
                        ok = True
                    else:
                        cells.append(MISSING)
                if ok:
                    rows.append(tuple(cells))
            if rows:
                headers = ["Показатель"] + \
                    [f"{y}" for y in self._years_desc()]
                fmt = "#,##0.0" if title == "Оборачиваемость, дней" \
                    else "0.00"
                ranges[title], _, _ = put_block(
                    f"{title}, значение по годам", headers, rows,
                    cell_fmt=fmt)

        for col, w in (("A", 42), ("B", 16), ("C", 16)):
            ws.column_dimensions[col].width = w
        return ranges

    # --- Диаграммы Excel ----------------------------------------------------- #
    def _add_excel_charts(self, wb, block_ranges):
        from openpyxl.chart import BarChart, PieChart, Reference
        from openpyxl.chart.series import DataPoint
        from openpyxl.chart.label import DataLabelList
        from openpyxl.chart.data_source import AxDataSource, StrRef
        import openpyxl.chart.data_source as _ds
        import openpyxl.chart.series as _ss
        _StrData = _ds.StrData
        _StrVal = _ds.StrData.pt.expected_type  # openpyxl.chart.data_source.StrVal
        _Series = _ss.Series

        ws = wb[CHARTS_SHEET] if CHARTS_SHEET in wb.sheetnames \
            else wb.create_sheet(CHARTS_SHEET)
        from openpyxl.styles import Font as _Font
        ws.cell(row=1, column=1,
                value="Диаграммы построены на области «Данные для диаграмм» "
                      "(редактируемые: тип диаграмм, цвета, метки данных, "
                      "оси настраиваются средствами Excel).")
        ws.cell(row=1, column=1).font = _Font(bold=True)

        data_ws = wb[DATA_SHEET]
        years_all = self._years_desc()
        _base = [c.lstrip("#") for c in
                 ((self.style or {}).get("series") or _DEFAULT_SERIES)]
        series_colors = [_base[i % len(_base)]
                         for i in range(max(len(years_all), 1))]
        green = ((self.style or {}).get("dynamics") or _DEFAULT_DYNAMICS)

        def ref(name, col, header_row, last_row):
            return Reference(data_ws, min_col=col, max_col=col,
                             min_row=header_row + 1, max_row=last_row)

        def cat_ref(header_row, last_row):
            """Категории (статьи/показатели) как текстовая ссылка со
            strCache: числовая (numRef) без кеша оставляла подписи осей
            пустыми (проверено экспортом через Excel COM)."""
            pts = []
            names = []
            for r in range(header_row + 1, last_row + 1):
                txt = data_ws.cell(row=r, column=1).value
                if txt is None or str(txt) == MISSING:
                    txt = ""
                text = str(txt)
                p = _StrVal(idx=r - header_row - 1)
                p.v = text
                pts.append(p)
                names.append(text)
            sd = _StrData(ptCount=len(pts), pt=pts)
            sr = StrRef(f=f"'{DATA_SHEET}'!$A${header_row + 1}:"
                        f"$A${last_row}", strCache=sd)
            return AxDataSource(strRef=sr)

        def rich_sz(sz: int):
            """TextProperties c шрифтом заданного размера (пункты) — для
            меток данных и подписей осей."""
            from openpyxl.chart.text import RichText
            from openpyxl.drawing.text import (Paragraph,
                                               ParagraphProperties,
                                               CharacterProperties,
                                               RichTextProperties)
            cp = CharacterProperties(sz=int(sz * 100))
            return RichText(
                bodyPr=RichTextProperties(),
                p=[Paragraph(pPr=ParagraphProperties(defRPr=cp),
                             endParaRPr=cp)])

        def set_title(chart, text):
            """Заголовок диаграммы — 14 пт на всех диаграммах."""
            from openpyxl.drawing.text import (ParagraphProperties,
                                               CharacterProperties)
            chart.title = text
            cp = CharacterProperties(sz=1400, b=True)
            para = chart.title.tx.rich.p[0]
            para.pPr = ParagraphProperties(defRPr=cp)
            for run in (para.r or []):
                run.rPr = cp

        def add_labels(chart, num_fmt, show_cat=False, show_pct=False,
                       show_ser=False):
            """Метки данных (решение 30.10.2026): значения и/или доли,
            названия категорий/рядов; вне области ряда (outEnd); шрифт 8 пт."""
            dl = DataLabelList(showVal=not show_pct,
                               showCatName=show_cat, showSerName=show_ser,
                               showPercent=show_pct, showLegendKey=False,
                               showBubbleSize=False)
            dl.numFmt = num_fmt
            dl.dLblPos = "outEnd"
            dl.txPr = rich_sz(8)
            chart.dataLabels = dl

        def apply_series_colors(chart, colors):
            for i, ser in enumerate(chart.series):
                ser.graphicalProperties.solidFill = colors[
                    i % len(colors)]
        def pie(name, anchor):
            if name not in block_ranges:
                return
            col0, hrow, col1, last = block_ranges[name]
            pie_chart = PieChart()
            # суффикс « (круговая)» — служебный (имя блока данных)
            set_title(pie_chart,
                      f"{name[:-len(' (круговая)')]}, млн руб.")
            pie_chart.height, pie_chart.width = 8.5, 13
            data = Reference(data_ws, min_col=col0 + 1, max_col=col0 + 1,
                             min_row=hrow, max_row=last)
            pie_chart.add_data(data, titles_from_data=True)
            pie_chart.series[0].cat = cat_ref(hrow, last)
            pts = [DataPoint(idx=i)
                   for i in range(max(last - hrow, 1))]
            # Топ-3 — акцентные (салатовый/голубой/оранжевый), остальные —
            # серые оттенки; при отсутствии расчета — прежние палитры
            rank_colors = self._pie_colors.get(name) or _style_ramps(self.style)
            for i, pt in enumerate(pts):
                pt.graphicalProperties.solidFill = \
                    rank_colors[i % len(rank_colors)].lstrip("#")
            pie_chart.series[0].data_points = pts
            # Названия статей в метках (легенда избыточна) + доля,
            # срезы <1% не показываются — блок «(круговая)» уже сводный
            add_labels(pie_chart, num_fmt="0.0%", show_cat=True,
                       show_pct=True)
            pie_chart.legend = None
            ws.add_chart(pie_chart, anchor)

        def cluster(name, anchor, horizontal=True, n_series=None,
                    color=None, title=None, num_fmt="#,##0.0",
                    show_ser=False, axis_fmt=None, axis_labels=True,
                    legend=None):
            """Линейчатая/столбчатая. legend=None: у линейчатых легенда
            убрана (решение 30.10.2026), у многорядных (по явному запросу)
            возвращается. axis_labels=False — у линейчатых «млн руб.» ось
            значений без подписей: числа читаются в метках данных.
            Все подписи и оси категорий — шрифт 8 пт."""
            if name not in block_ranges:
                return
            col0, hrow, col1, last = block_ranges[name]
            chart = BarChart()
            chart.type = "bar" if horizontal else "col"
            set_title(chart, title or f"{name}, млн руб.")
            chart.height, chart.width = 10, 16
            chart.gapWidth = 60
            chart.varyColors = False  # иначе Excel легенду по категориям
            chart.x_axis.delete = False
            chart.x_axis.tickLblPos = "nextTo"
            chart.x_axis.txPr = rich_sz(8)  # названия категорий — 8 пт
            chart.y_axis.delete = False
            chart.y_axis.tickLblPos = "nextTo"
            chart.y_axis.txPr = rich_sz(8)
            if legend is None:
                legend = not horizontal  # линейчатые — без легенды
            if not legend:
                chart.legend = None
            n_series = n_series if n_series is not None else (col1 - col0)
            cat = cat_ref(hrow, last)
            for j in range(n_series):
                data = Reference(data_ws, min_col=col0 + 1 + j,
                                 max_col=col0 + 1 + j, min_row=hrow,
                                 max_row=last)
                from openpyxl.chart import Series
                ser = Series(data, title_from_data=True)
                ser.cat = cat
                ser.invertIfNegative = False  # иначе отрицательные столбцы белые
                fill_hex = (color if color else series_colors[j]).lstrip("#")
                ser.graphicalProperties.solidFill = fill_hex
                chart.series.append(ser)
            add_labels(chart, num_fmt=num_fmt, show_ser=show_ser)
            if axis_fmt:
                chart.y_axis.number_format = axis_fmt
            if not axis_labels:
                # линейчатые «млн руб.»: ось значений скрыта — числа читаются
                # в метках данных (как на панели приложения)
                chart.y_axis.delete = True
            chart.y_axis.majorGridlines = None
            ws.add_chart(chart, anchor)

        pie("Структура активов (круговая)", "A4")
        pie("Структура капитала (круговая)", "I4")
        cluster("Динамика изменений статей Баланса", "A22",
                horizontal=True, n_series=1, color=green,
                title="Динамика изменений статей Баланса, млн руб.",
                axis_labels=False)
        # легенда годов возвращена по явному запросу (30.10.2026);
        # годы из меток данных убраны
        cluster("Финансовые результаты", "I22", horizontal=True,
                title="Финансовые результаты, млн руб.",
                axis_labels=False, legend=True)
        cluster("Ликвидность", "A42", horizontal=False,
                title="Ликвидность, значение по годам", num_fmt="0.00")
        cluster("Финансовая устойчивость", "I42", horizontal=False,
                title="Финансовая устойчивость, значение по годам",
                num_fmt="0.00")
        cluster("Рентабельность", "A62", horizontal=False,
                title="Рентабельность, %, значение по годам",
                num_fmt="0.00")
        cluster("Оборачиваемость, дней", "I62", horizontal=False,
                title="Оборачиваемость, дней, значение по годам",
                num_fmt="#,##0.0")

    # =====================================================================
    # Word (FR-07, AC-05) — текст по шаблону reference/brief_rezult.pdf
    # =====================================================================
    def generate_docx(self, path: str, xlsx_path: Optional[str] = None) -> None:
        """xlsx_path — путь к explonation.xlsx того же пакета: диаграммы
        копируются из него в Word как редактируемые OLE-объекты (COM).
        Без Excel — PNG-графики в той же палитре, что на панели.
        Режим вставки сохраняется в self.docx_chart_mode ("excel"/"png")."""
        self._build_doc(path)
        pasted = False
        if xlsx_path and os.path.exists(xlsx_path):
            pasted = self._word_insert_charts(path, xlsx_path)
        self.docx_chart_mode = "excel" if pasted else "png"
        if not pasted:
            self._build_doc(path, png_charts=True)

    def _build_doc(self, path: str, png_charts: bool = False) -> None:
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import datetime

        doc = Document()
        accent_hex = ((self.style or {}).get("series") or
                      ["2A78D6"])[0].lstrip("#")
        accent = RGBColor(int(accent_hex[0:2], 16), int(accent_hex[2:4], 16),
                          int(accent_hex[4:6], 16))
        year = self._year_label()
        self.content.set_value_getter(self._v)

        def heading(text, level=1):
            h = doc.add_heading(text, level=level)
            for run in h.runs:
                run.font.color.rgb = accent

        def bullet(text):
            doc.add_paragraph(text, style="List Bullet")

        # --- Шапка (по шаблону) ------------------------------------------- #
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("Заключение Ревизионной комиссии")
        run.bold = True
        run.font.size = Pt(14)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(f"{ANON_MARK} (далее «Общество»)").bold = True
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(
            f"по результатам проверки финансово-хозяйственной деятельности "
            f"за {year} год.")
        run.bold = True
        run.font.size = Pt(13)
        doc.add_paragraph(f"г. {ANON_MARK}")
        doc.add_paragraph(
            "Дата подписания: [требует уточнения]. Заключение сформировано "
            f"{datetime.date.today().strftime('%d.%m.%Y')}.")
        doc.add_paragraph(
            "Примечание: в документе удалены наименование организации и "
            "персональные реквизиты (BR-03).")

        # --- 1. Состав РК и основание --------------------------------------- #
        heading(f"1. Состав Ревизионной комиссии и основание проведения "
                f"ревизионной проверки за {year}г.")
        doc.add_paragraph(
            f"Ревизионная комиссия (далее «РК») в составе: {ANON_MARK}, "
            f"утвержденная на годовом общем собрании акционеров (реквизиты "
            f"решения — требует уточнения), провела проверку "
            f"финансово-хозяйственной деятельности Общества за {year} г. "
            f"на основании годовой отчетности Общества, в состав которого "
            f"входят:")
        bullet("бухгалтерский баланс, отчет о финансовых результатах, "
               "отчет об изменениях капитала, отчет о движении денежных "
               "средств, пояснения к бухгалтерскому балансу и отчету о "
               f"финансовых результатах за {year} год;")
        bullet(f"годовой отчет Общества за {year} год;")
        bullet(f"аудиторское заключение [ДАННЫЕ УДАЛЕНЫ] (далее «Аудитор») "
               f"от [ДАННЫЕ УДАЛЕНЫ] (требует уточнения);")
        bullet("документов, предоставленных Обществом на основании базового "
               "запроса Ревизионной комиссии, в том числе первичные "
               "документы и договоры, являющимися основанием для отражения "
               "операций в бухгалтерской (финансовой) отчетности Общества "
               f"за {year}г.")
        doc.add_paragraph(
            f"Заключение сформировано по результатам проверки Общества за "
            f"{year} год на основании Закона №208-ФЗ от 24.11.1995 г. "
            f"«Об акционерных обществах», Устава Общества и Положения о "
            f"Ревизионной комиссии Общества (новая редакция) от [ДАННЫЕ "
            f"УДАЛЕНЫ], основной целью которого является подтверждение "
            f"достоверности данных годового отчета, годовой бухгалтерской "
            f"(финансовой) отчетности Общества, при этом ответственность "
            f"за подготовку и достоверное представление данных в годовом "
            f"отчете и бухгалтерской (финансовой) отчетности Общества в "
            f"соответствии с действующим законодательством РФ и правилами "
            f"составления бухгалтерской (финансовой) отчетности в РФ несет "
            f"руководство и лица Общества, ответственные за корпоративное "
            f"управление и за составление бухгалтерской (финансовой) "
            f"отчетности.")

        # --- 2. Объем проверки и основные показатели ------------------------ #
        heading("2. Объем проверки и основные показатели "
                "финансово-хозяйственной деятельности Общества.")
        doc.add_paragraph("2.1. На основании представленных документов:")
        bullet("проведен выборочный анализ: системы управления Обществом, "
               "текущего и стратегического планирования, учетной политики "
               "для целей бухгалтерского учета и налогообложения, "
               "формирования основных статей бухгалтерского баланса "
               "(основные средства, запасы, дебиторская и кредиторская "
               "задолженность, кредиты и займы, собственный капитал) и "
               "финансового результата в части формирования доходов и "
               "расходов отчета о финансовых результатах согласно учетной "
               "политики Общества;")
        bullet("проведен анализ динамики основных статей баланса и отчета о "
               "финансовых результатах, показателей ликвидности, "
               "финансовой устойчивости, эффективности, показателей "
               "прогнозирования банкротства и финансового кризиса (модель "
               "Альтмана) (Приложение 1 к настоящему Заключению).")
        doc.add_paragraph(
            f"2.2. Основные показатели финансово-хозяйственной деятельности "
            f"в соответствии с бухгалтерской (финансовой) отчетностью за "
            f"{year} и в динамике по годам представлены в таблице:")

        years_all = self._years_desc()
        rows_spec = [
            ("Баланс", "1100", "1. Внеоборотные активы"),
            ("Баланс", "1200", "2. Оборотные активы"),
            ("Баланс", "1600", "3. Активы — всего"),
            ("ОФР", "2110", "4. Выручка"),
            ("ОФР", "2120", "5. Себестоимость"),
            ("ОФР", "2100", "6. Валовая прибыль (убыток)"),
            ("ОФР", "2210", "7. Коммерческие расходы"),
            ("ОФР", "2220", "8. Управленческие расходы"),
            ("ОФР", "2200", "9. Прибыль (убыток) от продаж"),
            ("ОФР", "2330", "10. Проценты к уплате"),
            ("ОФР", "2340", "11. Прочие доходы"),
            ("ОФР", "2350", "12. Прочие расходы"),
            ("ОФР", "2300", "13. Прибыль (убыток) до налогообложения"),
            ("ОФР", "2400", "14. Чистая прибыль (убыток)"),
            ("ОДДС", "4100", "15. Сальдо денежного потока от текущих "
                             "операций"),
            ("ОДДС", "4400", "16. Сальдо денежного потока за период"),
        ]
        table = doc.add_table(rows=1, cols=2 + len(years_all))
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "Показатели, тыс. руб."
        hdr[1].text = "Источник"
        for j, y in enumerate(years_all):
            hdr[2 + j].text = f"{y}"
        for form, code, label in rows_spec:
            cells = table.add_row().cells
            cells[0].text = label
            cells[1].text = f"{form}, стр. {code}"
            for j, y in enumerate(years_all):
                v = self._v(form, code, y)
                cells[2 + j].text = ru_num(v, 0) if v is not None else MISSING
        for row_ in table.rows:
            for c in row_.cells:
                for par in c.paragraphs:
                    for run in par.runs:
                        run.font.size = Pt(9)

        # --- 3. Выводы ------------------------------------------------------- #
        heading("3. Выводы:")
        doc.add_paragraph(
            f"По результатам проведенной проверки за {year} год на основании "
            f"представленных Обществом документов и аудиторского заключения "
            f"Аудитора за {year} г., Ревизионная комиссия пришла к "
            f"следующим выводам:")
        concls = self.content.conclusions()
        for i, concl in enumerate(concls, 1):
            doc.add_paragraph(f"3.{i}. {concl}")
        doc.add_paragraph(
            f"3.7. {self.content.passive_influence()}")

        doc.add_paragraph("3.8. Таблица «Риск банкротства»:")
        self._bankruptcy_table(doc, years_all)
        doc.add_paragraph(
            "3.9. Статусы показателей рассчитаны сравнением с "
            "рекомендуемым диапазоном методологии (BR-01); отсутствующие "
            f"данные отмечены как «{MISSING}».")

        # --- 4. Рекомендации -------------------------------------------------- #
        heading("4. Рекомендации:")
        for i, rec in enumerate(self.content.concrete_recommendations(), 1):
            doc.add_paragraph(f"4.{i}. {rec}")

        # --- Концовка (по шаблону) --------------------------------------------- #
        doc.add_paragraph(
            f"Настоящее заключение должно рассматриваться совместно со всеми "
            f"формами бухгалтерской (финансовой) отчетности на "
            f"31.12.{year} года, отчетом о финансовых результатах и "
            f"пояснениями к бухгалтерскому балансу за {year} год, годовым "
            f"отчетом и заключением Аудитора за {year} год.")
        doc.add_paragraph(
            "Заключение составлено в 2-х экземплярах.")
        doc.add_paragraph(
            f"Члены Ревизионной Комиссии: {ANON_MARK}.")
        doc.add_paragraph(
            f"Председатель ревизионной комиссии: {ANON_MARK}.")

        # --- Приложение 1: диаграммы ------------------------------------------- #
        heading("Приложение 1 к Заключению РК")
        for name in CHART_ORDER:
            doc.add_paragraph(name)
            if png_charts:
                img = self._ensure_charts().get(name)
                if img and os.path.exists(img):
                    doc.add_picture(img, width=Inches(6.0))
                else:
                    doc.add_paragraph(f"Диаграмма «{name}» недоступна.")
            else:
                doc.add_paragraph(f"[ДИАГРАММА: {name}]")
        doc.save(path)

    def _bankruptcy_table(self, doc, years_all) -> None:
        from docx.shared import Pt
        rows_spec = [
            ("Z-счет Альтмана (двухфакторная модель)",
             "Z < 0 — вероятность банкротства < 50%"),
            ("Z-счет Альтмана (пятифакторная модель)",
             "> 2,9 — зеленая; 1,8–2,9 — серая; < 1,8 — красная"),
            ("Модель Таффлера",
             "> 0,3 — устойчив; 0,2–0,3 — серая; < 0,2 — риск"),
        ]
        by_name = self.content.by_name
        table = doc.add_table(rows=1, cols=2 + len(years_all))
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "Модель"
        hdr[1].text = "Зоны риска"
        for j, y in enumerate(years_all):
            hdr[2 + j].text = f"{y} г."
        for name, zone_desc in rows_spec:
            r = by_name.get(name)
            cells = table.add_row().cells
            cells[0].text = name
            cells[1].text = zone_desc
            for j, y in enumerate(years_all):
                v = r["values"].get(y) if r else None
                cells[2 + j].text = ru_num(v, 2) if v is not None else MISSING
            # Статус последним столбцом не добавляем — зоны в описании
            if r is not None:
                extra = table.add_row().cells
                extra[0].text = ""
                extra[1].text = f"Статус ({status_word(r['status'])})"
                for j in range(len(years_all)):
                    extra[2 + j].text = ""
        for row_ in table.rows:
            for c in row_.cells:
                for par in c.paragraphs:
                    for run in par.runs:
                        run.font.size = Pt(9)

    def _word_insert_charts(self, docx_path: str, xlsx_path: str) -> bool:
        """Копирование диаграмм из explonation.xlsx в rezult.docx как
        редактируемых OLE-объектов (требуется установленный Excel и Word).
        Возвращает True, если все вставки прошли успешно. При ошибке ее
        текст кладется в self.insert_error (для диагностики в UI)."""
        self.insert_error = None
        try:
            import pythoncom
            import win32com.client
        except ImportError:
            self.insert_error = "pywin32 не установлен"
            return False
        excel = word = None
        wb = None
        doc = None
        ok = 0
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            # ReadOnly: файл может быть открыт пользователем в Excel
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path), 0, True)
            sheet = wb.Worksheets(CHARTS_SHEET)
            n_charts = sheet.ChartObjects().Count
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(docx_path))
            planned = min(n_charts, len(CHART_ORDER))
            import time as _time
            for i, name in enumerate(CHART_ORDER[:planned], start=1):
                obj = sheet.ChartObjects(i)
                sel = word.Selection
                sel.HomeKey(Unit=6)                     # wdStory
                f = sel.Find
                f.ClearFormatting()
                if not f.Execute(FindText=f"[ДИАГРАММА: {name}]",
                                 MatchCase=False, MatchWholeWord=False,
                                 MatchWildcards=False, Forward=True,
                                 Wrap=0, Format=False):
                    self.insert_error = f"placeholder не найден: {name}"
                    break
                sel.Text = ""                            # удалить placeholder
                # Копирование Excel в буфер асинхронно: пауза и повторы
                pasted = False
                for attempt in range(4):
                    obj.Copy()
                    _time.sleep(1.0 + attempt * 0.8)
                    try:
                        sel.PasteSpecial(0)              # wdPasteOLEObject
                        pasted = True
                        break
                    except Exception:
                        continue
                if not pasted:
                    self.insert_error = f"буфер обмена не отдал диаграмму: {name}"
                    break
                try:                                     # ширина в документе
                    shape = word.Selection.InlineShapes(1)
                    shape.LockAspectRatio = -1
                    shape.Width = word.CentimetersToPoints(16)
                except Exception:
                    pass
                ok += 1
            if ok != planned and self.insert_error is None:
                self.insert_error = f"вставлено {ok} из {planned}"
            doc.Save()
            return ok == planned and planned > 0
        except Exception as e:
            import traceback
            self.insert_error = traceback.format_exc(limit=3)
            return False
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            try:
                if wb is not None:
                    wb.Close(False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    # =====================================================================
    # PowerPoint (FR-08, AC-09) — нативные редактируемые диаграммы
    # =====================================================================
    def generate_pptx(self, path: str) -> None:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION,
                                     XL_LABEL_POSITION)

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        ink = RGBColor(0x0B, 0x0B, 0x0B)
        muted = RGBColor(0x52, 0x51, 0x4E)
        years_all = self._years_desc()
        year = self._year_label()
        self.content.set_value_getter(self._v)
        slide_concl = self.content.slide_conclusions()
        chart_conc = self.content.conclusions()

        series_hex = (self.style or {}).get("series",
                                            ["2A78D6", "EB6834", "8A877F"])

        def blank():
            return prs.slides.add_slide(prs.slide_layouts[6])

        def txt(slide, left, top, width, height, text, size=18, bold=False,
                color=ink):
            box = slide.shapes.add_textbox(Inches(left), Inches(top),
                                           Inches(width), Inches(height))
            tf = box.text_frame
            tf.word_wrap = True
            for i, line in enumerate(text.split("\n")):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                r = p.add_run()
                r.text = line
                r.font.size = Pt(size)
                r.font.bold = bold
                r.font.color.rgb = color
            return box

        def native_chart(slide, chart_type, cats, series, left, top, width,
                         height, colors, num_fmt="0.00", show_pct=False,
                         show_cat=False, show_ser=False, single_color=None,
                         title=None, slice_colors=None, legend=True):
            """Редактируемая нативная диаграмма PowerPoint (не рисунок)."""
            data = CategoryChartData()
            data.categories = cats
            for label, vals in series:
                data.add_series(label, vals)
            gf = slide.shapes.add_chart(chart_type, Inches(left),
                                        Inches(top), Inches(width),
                                        Inches(height), data)
            ch = gf.chart
            if title:
                ch.has_title = True
                ch.chart_title.text_frame.text = title
                for r in ch.chart_title.text_frame.paragraphs[0].runs:
                    r.font.size = Pt(14)
            plot = ch.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = True
            if show_ser:
                dl.show_series_name = True
            dl.number_format = num_fmt
            dl.number_format_is_linked = False
            dl.font.size = Pt(8)
            if chart_type == XL_CHART_TYPE.PIE:
                dl.show_value = False
                if show_pct:
                    dl.show_percentage = True
                if show_cat:
                    dl.show_category_name = True
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
                plot.vary_by_categories = True
                if slice_colors:
                    for i, pt in enumerate(ch.series[0].points):
                        pt.format.fill.solid()
                        pt.format.fill.fore_color.rgb = \
                            RGBColor.from_string(
                                str(slice_colors[i]).lstrip("#"))
            else:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
                for i, ser in enumerate(ch.series):
                    hexc = single_color if single_color else colors[
                        i % len(colors)]
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.rgb = \
                        RGBColor.from_string(hexc.lstrip("#"))
                    # без этого отрицательные столбцы рендерятся белыми
                    _ser_no_invert(ser)
                # названия категорий/показателей по оси — 8 пт
                try:
                    ch.category_axis.tick_labels.font.size = Pt(8)
                except (KeyError, AttributeError, TypeError):
                    pass  # ось отсутствует на этом типе диаграммы
                if legend:
                    ch.has_legend = True
                    ch.legend.position = XL_LEGEND_POSITION.TOP
                    ch.legend.include_in_layout = False
                    ch.legend.font.size = Pt(10)
            return ch

        def concluder(title):
            """Вывод слайда (2–4 предложения, выделенный шрифт)."""
            return slide_concl.get(title, "")

        def slide_chart(title, cats, series, colors, num_fmt, **kw):
            s = blank()
            txt(s, 0.6, 0.25, 12.2, 0.6, title, size=26, bold=True)
            native_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cats, series,
                         2.0, 0.95, 9.3, 5.0, colors, num_fmt=num_fmt,
                         title=None, **kw)
            concl = concluder(title)
            if concl:
                txt(s, 0.6, 6.05, 12.2, 1.3, concl, size=15, bold=True)
            return s

        # Слайд 1 — титул
        s = blank()
        txt(s, 0.8, 2.4, 11.7, 1.2, "Результаты финансового анализа",
            size=40, bold=True)
        txt(s, 0.8, 3.6, 11.7, 1.0,
            f"Заключение по данным бухгалтерской отчетности за {year} год\n"
            "Анонимизированный пакет отчетов (BR-03)", size=16, color=muted)

        # Слайд 2 — основные показатели (таблица)
        s = blank()
        txt(s, 0.6, 0.3, 12.2, 0.6,
            "Основные показатели финансово-хозяйственной деятельности, "
            "млн руб.", size=22, bold=True)
        mt = self.results["main_table"]
        n_rows = len(mt) + 1
        tbl = s.shapes.add_table(n_rows, 2 + len(years_all),
                                 Inches(0.6), Inches(1.05),
                                 Inches(12.1),
                                 Inches(min(0.34 * n_rows, 5.2))).table
        tbl.cell(0, 0).text = "Показатель"
        tbl.cell(0, 1).text = "Источник"
        for j, y in enumerate(years_all):
            tbl.cell(0, 2 + j).text = f"{y} г."
        for i, (_, r) in enumerate(mt.iterrows(), 1):
            tbl.cell(i, 0).text = str(r["Показатель"])
            tbl.cell(i, 1).text = str(r["Источник"])
            for j, y in enumerate(years_all):
                col = f"{y} г., млн руб."
                tbl.cell(i, 2 + j).text = str(r[col]) \
                    if col in mt.columns else MISSING
        for i in range(n_rows):
            for j in range(2 + len(years_all)):
                for p in tbl.cell(i, j).text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(10)
        txt(s, 0.6, 6.7, 12.2, 0.7, chart_conc[0] if chart_conc else "",
            size=14, bold=True)

        # Слайды KPI по годам (редактируемые гистограммы)
        by_name = self._by_name()
        for title, names_map in KPI_GROUPS.items():
            nd = KPI_KIND[title][0]
            cats, series = [], []
            rows_used = []
            for name, short in names_map.items():
                row = by_name.get(name)
                if row is None or all(row["values"].get(y) is None
                                      for y in years_all):
                    continue
                cats.append(short)
                rows_used.append(row)
            if not cats:
                continue
            series = [(str(y), [r["values"].get(y) for r in rows_used])
                      for y in years_all]
            slide_chart(title, cats, series, series_hex,
                        "0.00" if nd == 2 else "0")

        # Круговые структуры: статьи >=1% + «Прочие <1%»; топ-3 — акцентные
        # (салатовый/голубой/оранжевый, key стиля pie_top), остальные — серые

        def prep_pie(cd):
            val_col = [c for c in cd.columns
                       if str(c).startswith("Значение, млн")][0]
            raw = [(str(r["Статья"]), _num(r[val_col]))
                   for _, r in cd.iterrows()
                   if _num(r[val_col]) is not None and _num(r[val_col]) > 0]
            items = merge_small(raw)
            colors = accent_colors(items, self.style)
            return items, colors

        for name in ("Структура активов", "Структура капитала"):
            cd = self.results.get("chart_data", {}).get(name)
            s = blank()
            txt(s, 0.6, 0.25, 12.2, 0.6, name, size=26, bold=True)
            if cd is not None:
                items, slice_colors = prep_pie(cd)
                if items:
                    cats = [n for n, _ in items]
                    vals = [v for _, v in items]
                    native_chart(s, XL_CHART_TYPE.PIE, cats,
                                 [(year, vals)], 1.4, 0.95,
                                 10.0, 5.0, ["FFFFFF"], num_fmt="0.0%",
                                 show_pct=True, show_cat=True, legend=False,
                                 slice_colors=slice_colors,
                                 title=f"{name}, млн руб.")
            concl = concluder(name)
            if concl:
                txt(s, 0.6, 6.15, 12.2, 1.2, concl, size=15, bold=True)

        # Динамика статей Баланса (линейчатая)
        dyn = self.results.get("chart_data", {}).get(
            "Динамика изменений статей Баланса")
        if dyn is not None:
            items = [(str(r["Статья"]), _num(r["Изменение, млн руб."]))
                     for _, r in dyn.iterrows()]
            items = [(n, v) for n, v in items if v is not None]
            if items:
                s = blank()
                txt(s, 0.6, 0.25, 12.2, 0.6,
                    "Динамика изменений статей Баланса", size=26, bold=True)
                cats = [n for n, _ in items][::-1]
                vals = [v for _, v in items][::-1]
                green = (self.style or {}).get("dynamics", "359C56")
                native_chart(
                    s, XL_CHART_TYPE.BAR_CLUSTERED, cats,
                    [("Изменение, млн руб.", vals)], 3.4, 0.95, 9.6, 5.2,
                    ["359C56"], num_fmt="#,##0.0", single_color=green,
                    legend=False)
                concl = concluder("Динамика изменений статей Баланса")
                if concl:
                    txt(s, 0.6, 6.25, 12.2, 1.1, concl, size=14, bold=True)

        # Финансовые результаты (линейчатая, 2 ряда)
        fin = self.results.get("chart_data", {}).get("Финансовые результаты")
        if fin is not None:
            val_cols = [c for c in fin.columns
                        if str(c).startswith("Значение, млн")]
            cats = [str(r["Статья"]) for _, r in fin.iterrows()][::-1]
            series = []
            for col, y in zip(val_cols, years_all):
                series.append((str(y),
                               [_num(r[col]) for _, r in fin.iterrows()]
                               [::-1]))
            if cats:
                s = blank()
                txt(s, 0.6, 0.25, 12.2, 0.6, "Финансовые результаты",
                    size=26, bold=True)
                native_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, cats, series,
                             2.4, 0.95, 10.2, 5.2, series_hex, num_fmt="#,##0.0")
                concl = concluder("Финансовые результаты")
                if concl:
                    txt(s, 0.6, 6.25, 12.2, 1.1, concl, size=14, bold=True)

        # Риски банкротства (таблица + вывод)
        s = blank()
        txt(s, 0.6, 0.25, 12.2, 0.6, "Риски банкротства", size=26, bold=True)
        by_name = self._by_name()
        rows_spec = [
            ("Z-счет Альтмана (двухфакторная модель)",
             "Z < 0 — вероятность банкротства < 50%"),
            ("Z-счет Альтмана (пятифакторная модель)",
             "> 2,9 — зеленая; 1,8–2,9 — серая; < 1,8 — красная"),
            ("Модель Таффлера",
             "> 0,3 — устойчив; 0,2–0,3 — серая; < 0,2 — риск"),
        ]
        tbl = s.shapes.add_table(4, 2 + len(years_all), Inches(0.6),
                                 Inches(1.0), Inches(12.1),
                                 Inches(2.0)).table
        tbl.cell(0, 0).text = "Модель"
        tbl.cell(0, 1).text = "Зоны риска"
        for j, y in enumerate(years_all):
            tbl.cell(0, 2 + j).text = f"{y} г."
        for i, (name, zone_desc) in enumerate(rows_spec, 1):
            r = by_name.get(name)
            tbl.cell(i, 0).text = name
            tbl.cell(i, 1).text = zone_desc
            for j, y in enumerate(years_all):
                v = r["values"].get(y) if r else None
                tbl.cell(i, 2 + j).text = ru_num(v, 2) if v is not None \
                    else MISSING
        for i in range(4):
            for j in range(2 + len(years_all)):
                for p in tbl.cell(i, j).text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(11)
        concl = concluder("Риски банкротства") or \
            (chart_conc[5] if len(chart_conc) > 5 else "")
        if concl:
            txt(s, 0.6, 5.7, 12.2, 1.6, concl, size=15, bold=True)

        # Финальный слайд: Выводы и Рекомендации в 2 колонки
        s = blank()
        txt(s, 0.6, 0.3, 12.2, 0.7, "Выводы и рекомендации", size=26,
            bold=True)
        concl_text = "\n".join("• " + t for t in
                               self.content.qualitative_conclusions())
        recs = self.content.concrete_recommendations()
        recs_text = "\n".join("• " + t for t in recs)
        txt(s, 0.6, 1.2, 6.0, 5.9, "ВЫВОДЫ\n" + concl_text, size=15,
            bold=True)
        txt(s, 6.9, 1.2, 6.0, 5.9, "РЕКОМЕНДАЦИИ\n" + recs_text, size=15,
            bold=True)

        prs.save(path)

    def _by_name(self) -> Dict[str, dict]:
        return {r["name"]: r for r in self.results["kpi"]}

    # =====================================================================
    # Сборка пакета
    # =====================================================================
    def generate_package(self, out_dir: str) -> Dict[str, str]:
        os.makedirs(out_dir, exist_ok=True)
        paths = {
            "excel": os.path.join(out_dir, "explonation.xlsx"),
            "docx": os.path.join(out_dir, "rezult.docx"),
            "pptx": os.path.join(out_dir, "presentation.pptx"),
        }
        self.generate_excel(paths["excel"])
        self.generate_docx(paths["docx"], xlsx_path=paths["excel"])
        self.generate_pptx(paths["pptx"])
        self.cleanup()
        return paths

    def cleanup(self) -> None:
        if self._charts_dir:
            shutil.rmtree(self._charts_dir, ignore_errors=True)
            self._charts_dir = None


def default_out_dir() -> str:
    """Папка output/ на уровне проекта (вне src/)."""
    return os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "output"))
