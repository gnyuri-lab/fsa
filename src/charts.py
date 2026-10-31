"""Диаграммы аналитической панели (Этап 3) — Altair/Vega-Lite.

Палитры валидированы скриптом dataviz validate_palette.js:
- ряды по годам: #2a78d6 (голубой), #eb6834 (оранжевый); #8a877f (серый 2023) —
  намеренно обесцвеченный ряд по требованию брифа (компенсируется метками и таблицей);
- светло-зеленый оттенки (оборотные активы, капитал):
  #66b57f,#3e9e5c,#2b8849,#1e7139,#155c2c,#0e4821 — ordinal, все проверки PASS;
- серые оттенки (внеоборотные активы, обязательства):
  #aaa89f,#89877f,#6b6963,#4f4e48,#383733,#242321 — ordinal, все проверки PASS;
- фон диаграмм светло-серый: #f4f3f0 (по брифу).
"""
import re
from typing import Any, Dict, List, Optional

import altair as alt
import pandas as pd

# --- Базовый стиль (FR-11) -------------------------------------------------
BASE_STYLE = {
    "series": ["#2a78d6", "#eb6834", "#8a877f"],   # 2025 (голубой), 2024, 2023
    "series_names": [],                            # ["2025", "2024", "2023"]
    "pie_top": ["#a8e05f", "#2a78d6", "#eb6834"],  # салатовый, голубой, оранжевый —
                                                   # топ-3 сектора круговых (30.10.2026)
    "green_ramp": ["#66b57f", "#3e9e5c", "#2b8849", "#1e7139", "#155c2c", "#0e4821"],
    "gray_ramp": ["#aaa89f", "#89877f", "#6b6963", "#4f4e48", "#383733", "#242321"],
    "dynamics": "#359c56",                          # светло-зеленый ряд динамики
    "surface": "#f4f3f0",                           # фон светло-серый
    "grid_lines": "#e1e0d9",
    "ink": "#0b0b0b",
    "muted": "#52514e",
}

_HEX = re.compile(r"^#[0-9a-fA-F]{6}$")

SMALL_LABEL = "Прочие <1%"


def merge_small(items, threshold: float = 0.01):
    """Объединяет статьи с долей < 1% в «Прочие <1%» (решение пользователя
    30.10.2026: мелкие статьи учитываются одной категорией, а не скрываются).

    items: список кортежей (имя, значение) с положительными значениями.
    Возвращает список [(имя, значение)]: крупные статьи в исходном
    порядке + «Прочие <1%» в конце (если такие есть)."""
    total = sum(v for _, v in items if v > 0)
    kept, merged = [], 0.0
    for name, v in items:
        if v <= 0:
            continue
        if total and v / total < threshold:
            merged += v
        else:
            kept.append((name, v))
    if merged > 0:
        kept.append((SMALL_LABEL, merged))
    return kept


def accent_colors(items, style, top: int = 3):
    """Цвета секторов круговых (решение 30.10.2026): топ-3 по значению —
    акцентные (pie_top: салатовый, голубой, оранжевый), остальные —
    оттенки gray_ramp; «Прочие <1%» всегда серая.

    items: список кортежей (имя, значение), порядок = порядок секторов.
    Возвращает список hex-цветов, выровненный по items."""
    names = [i[0] for i in items]
    accents = (style or {}).get("pie_top") or BASE_STYLE["pie_top"]
    gray = (style or {}).get("gray_ramp") or BASE_STYLE["gray_ramp"]
    # индексы топ-N по величине, в порядке убывания значения
    idx_sorted = sorted([i for i, n in enumerate(names) if n != SMALL_LABEL],
                        key=lambda i: items[i][1], reverse=True)[:top]
    accent_of = {idx: r for r, idx in enumerate(idx_sorted)}
    colors, gi = [], 0
    for i, (name, _v) in enumerate(items):
        if i in accent_of:
            colors.append(accents[accent_of[i] % len(accents)])
        else:
            colors.append(gray[gi % len(gray)])
            gi += 1
    return colors


def merge_style(custom: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Сливает базовый стиль с корпоративным (JSON-шаблон, ключи при валидных hex)."""
    style = {k: (list(v) if isinstance(v, list) else v) for k, v in BASE_STYLE.items()}
    if not custom:
        return style
    for key, value in custom.items():
        if key not in style:
            continue
        if isinstance(value, list):
            valid = [v for v in value if isinstance(v, str) and _HEX.match(v)]
            if valid:
                style[key] = valid
        elif isinstance(value, str) and _HEX.match(value):
            style[key] = value
    return style


# --- общие настройки -------------------------------------------------------
def _font(style) -> str:
    return "system-ui, -apple-system, Segoe UI, sans-serif"


def _finish(chart: alt.Chart, style, title: Optional[str] = None, height=None):
    """Общие chrome-настройки: фон светло-серый, шрифты, без рамки."""
    out = chart
    if title:
        out = out.properties(
            title=alt.Title(text=title, anchor="start", font=_font(style),
                            fontSize=14, color=style["ink"]))
    if height:
        out = out.properties(height=height)
    return out.configure_view(
        fill=style["surface"], stroke=None
    ).configure_title(
        font=_font(style), fontSize=13, color=style["ink"]
            ).configure_legend(labelFont=_font(style), labelColor=style["muted"],
                               titleFont=_font(style), titleColor=style["muted"])


def _year_colors(style, years: List[int]) -> Dict[int, str]:
    """Цвет ряда по году: первый из years -> series[0] (голубой) и т.д."""
    return {year: style["series"][i % len(style["series"])]
            for i, year in enumerate(years)}


# --- 1. Круговая вторичная линейчатая (структуры) -----------------------------------
def structure_chart(title: str, items: pd.DataFrame, style) -> alt.Chart:
    """items: столбцы [Статья, Значение (млн), Цвет].
    Классический вид «круговая вторичная линейчатая» (как в шаблоне
    reference/brief_rezult.pdf): круговая диаграмма с долями на секторах
    (>=10%), сектор «Другие категории» и справа вторичный столбик с выноской
    статей с долей <10% (имя + доля). Одна композитная диаграмма."""
    import math
    import textwrap

    PIE_W, PIE_H, PIE_R = 360, 300, 122
    STRIP_W = 90
    OUT_W = 210

    df = items[items["Значение"] > 0].copy()
    total = df["Значение"].sum()
    if not len(df) or not total:
        return _finish(alt.Chart().mark_text(
            text=["Нет данных для диаграммы"], size=12, color=style["muted"]
        ).properties(width=500, height=60), style)
    df["Доля"] = df["Значение"] / total * 100

    main = df[df["Доля"] >= 10].sort_values(
        "Значение", ascending=False).reset_index(drop=True)
    sec = df[df["Доля"] < 10].sort_values(
        "Значение", ascending=False).reset_index(drop=True)

    # --- Круговая (сектор «Другие» при наличии статей <10%) ---
    pie_df = main[["Статья", "Значение", "Цвет"]].copy()
    if len(sec):
        pie_df.loc[len(pie_df)] = {"Статья": "Другие категории",
                                   "Значение": float(sec["Значение"].sum()),
                                   "Цвет": style["gray_ramp"][1]}
    pie_df["Доля"] = pie_df["Значение"] / pie_df["Значение"].sum() * 100
    pie_df["cum"] = pie_df["Доля"].cumsum() - pie_df["Доля"]
    pie_df["mid"] = (pie_df["cum"] + pie_df["Доля"] / 2) / 100 * 2 * math.pi
    pie_df["lx"] = PIE_W / 2 + pie_df["mid"].apply(
        lambda t: math.sin(t) * PIE_R * 0.55)
    pie_df["ly"] = PIE_H / 2 - pie_df["mid"].apply(
        lambda t: math.cos(t) * PIE_R * 0.55)
    pie_df["Метка"] = pie_df["Статья"].map(
        lambda s: "\n".join(textwrap.wrap(str(s), 16))) + \
        "\n" + pie_df["Доля"].map(lambda d: f"{d:.1f}%".replace(".", ","))

    arc = alt.Chart(pie_df).mark_arc(radius=PIE_R).encode(
        theta=alt.Theta("Доля:Q", stack=True),
        color=alt.Color("Цвет:N", scale=None, legend=None),
        tooltip=[alt.Tooltip("Статья:N", title="Статья"),
                 alt.Tooltip("Значение:Q", title="млн руб.", format=",.1f"),
                 alt.Tooltip("Доля:Q", title="Доля, %", format=".1f")])
    pie_labels = alt.Chart(pie_df).mark_text(
        size=9, align="center", baseline="middle", lineBreak="\n",
        color=style["ink"]).encode(
        x=alt.X("lx:Q", scale=alt.Scale(domain=[0, PIE_W]), axis=None),
        y=alt.Y("ly:Q", scale=alt.Scale(domain=[0, PIE_H]), axis=None),
        text=alt.Text("Метка:N"))
    pie_part = (arc + pie_labels).properties(width=PIE_W, height=PIE_H)

    parts = [pie_part]
    if len(sec):
        # --- Вторичный столбик (узкая линейчатая) ---
        sec["Выноска"] = sec["Статья"].map(
            lambda s: "\n".join(textwrap.wrap(str(s), 22))) + " " + \
            sec["Доля"].map(lambda d: f"({d:.1f}%)".replace(".", ","))
        strip = alt.Chart(sec).mark_bar(cornerRadiusEnd=2).encode(
            x=alt.X("Значение:Q", axis=None,
                    scale=alt.Scale(
                        domain=[0, float(sec["Значение"].max() * 1.05)])),
            y=alt.Y("Статья:N", sort=list(sec["Статья"]), axis=None,
            scale=alt.Scale(paddingInner=0.04, paddingOuter=0.04)),
            color=alt.Color("Цвет:N", scale=None, legend=None),
            tooltip=[alt.Tooltip("Статья:N", title="Статья"),
                     alt.Tooltip("Значение:Q", title="млн руб.", format=",.1f"),
                     alt.Tooltip("Доля:Q", title="Доля, %", format=".1f")])
        strip_part = strip.properties(width=STRIP_W, height=PIE_H)

        # --- Выноска: статьи <10% (имя + доля) ---
        callout = alt.Chart(sec).mark_text(
            size=10, align="left", dx=6, baseline="middle", lineBreak="\n",
            color=style["ink"]).encode(
            y=alt.Y("Статья:N", sort=list(sec["Статья"]), axis=None),
            x=alt.value(4),
            text=alt.Text("Выноска:N"))
        callout_part = callout.properties(width=OUT_W, height=PIE_H)
        parts += [strip_part, callout_part]

    compound = alt.hconcat(*parts) if len(parts) > 1 else pie_part
    compound = compound.properties(title=alt.Title(
        text=title, anchor="middle", font=_font(style),
        fontSize=14, color=style["ink"], dy=-6))
    return _finish(compound, style)


# --- 2. Динамика изменений статей Баланса (линейчатая, один ряд) --------------------
def dynamics_bars(df: pd.DataFrame, style) -> alt.Chart:
    """df: столбцы [Код, Статья, Изм (число, млн), Метка]. Светло-зеленый ряд,
    метки около ряда, шкала значений и сетка отсутствуют."""
    data = df[df["Изм"].notna()].copy()
    if not len(data):
        return alt.Chart().mark_text(
            text=["Нет данных"], size=12, color=style["muted"]
        ).properties(width=500, height=60)
    base = alt.Chart(data).mark_bar(color=style["dynamics"], size=14,
                                    cornerRadiusEnd=3).encode(
        x=alt.X("Изм:Q", axis=None),
        y=alt.Y("Статья:N", sort=list(data["Статья"]), axis=alt.Axis(
            orient="right", labelColor=style["muted"], labelFont=_font(style),
            labelFontSize=10, labelLimit=240, title=None,
            domain=False, ticks=False)),
        tooltip=[alt.Tooltip("Статья:N", title="Статья"),
                 alt.Tooltip("Метка:N", title="Изменение")])
    # Условное выравнивание меток знакомым значением — двумя слоями
    pos = data[data["Изм"] > 0]
    neg = data[data["Изм"] <= 0]
    txt_pos = alt.Chart(pos).mark_text(fontSize=10, align="left", dx=5,
                                       color=style["muted"]).encode(
        x=alt.X("Изм:Q"), y=alt.Y("Статья:N", sort=list(data["Статья"])),
        text=alt.Text("Метка:N"))
    txt_neg = alt.Chart(neg).mark_text(fontSize=10, align="right", dx=-5,
                                       color=style["muted"]).encode(
        x=alt.X("Изм:Q"), y=alt.Y("Статья:N", sort=list(data["Статья"])),
        text=alt.Text("Метка:N"))
    chart = base + txt_pos + txt_neg
    return _finish(chart, style, title="Динамика изменений статей Баланса, млн руб.",
                   height=max(120, 30 * len(data) + 40))


# --- 3. Финансовые результаты (линейчатая с группировкой, 2 ряда) --------------------
def fin_results_bars(df: pd.DataFrame, style, year: int, prev_year: int) -> alt.Chart:
    """df: столбцы [Код, Статья, Значение (отч.), Значение (пред.)] в млн."""
    val_cols = [c for c in df.columns if str(c).startswith("Значение, млн")]
    rows = []
    for _, r in df.iterrows():
        for vc, y in zip(val_cols, [year, prev_year]):
            v = r[vc]
            if pd.isna(v) or v == "требует уточнения":
                continue
            rows.append({"Статья": r["Статья"], "Год": str(y),
                         "Значение": float(v), "Метка": f"{float(v):+,.1f}".replace(",", " ")})
    if not rows:
        return alt.Chart().mark_text(
            text=["Нет данных"], size=12, color=style["muted"]
        ).properties(width=500, height=60)
    data = pd.DataFrame(rows)
    order = list(df["Статья"])
    colors = {str(year): style["series"][0], str(prev_year): style["series"][1]}
    vmin, vmax = data["Значение"].min(), data["Значение"].max()
    domain = [min(0, vmin * 1.35), max(0, vmax * 1.35)]
    base = alt.Chart(data).mark_bar(size=12, cornerRadiusEnd=2).encode(
        x=alt.X("Значение:Q", scale=alt.Scale(domain=domain), axis=None),
        y=alt.Y("Статья:N", sort=order, axis=alt.Axis(
            labelColor=style["muted"], labelFont=_font(style), labelFontSize=10,
            labelLimit=240, title=None, domain=False, ticks=False)),
        yOffset="Год:N",
        color=alt.Color("Год:N", scale=alt.Scale(domain=list(colors.keys()),
                                                 range=list(colors.values())),
                        legend=alt.Legend(title="Год", orient="top",
                                          labelColor=style["muted"])),
        tooltip=[alt.Tooltip("Статья:N", title="Статья"),
                 alt.Tooltip("Год:N"),
                 alt.Tooltip("Значение:Q", title="млн руб.", format=",.1f")])
    # Условное выравнивание меток: двумя слоями (Altair не допускает
    # conditional align/dx в encode())
    pos = data[data["Значение"] > 0]
    neg = data[data["Значение"] <= 0]
    txt = alt.layer(
        alt.Chart(pos).mark_text(fontSize=9, align="left", dx=4,
                                 color=style["muted"]).encode(
            x=alt.X("Значение:Q"), y=alt.Y("Статья:N", sort=order),
            yOffset="Год:N", text=alt.Text("Метка:N")),
        alt.Chart(neg).mark_text(fontSize=9, align="right", dx=-4,
                                 color=style["muted"]).encode(
            x=alt.X("Значение:Q"), y=alt.Y("Статья:N", sort=order),
            yOffset="Год:N", text=alt.Text("Метка:N")))
    chart = base + txt
    return _finish(chart, style, title="Финансовые результаты, млн руб.",
                   height=max(120, 34 * len(order)))


# --- 4. Гистограммы с группировкой (ликвидность, рентабельность, оборачиваемость) ---
def grouped_histogram(df: pd.DataFrame, style, title: str,
                      value_title: str) -> alt.Chart:
    """df (long): столбцы [Показатель, Год, Значение, Метка]. Ряды по годам:
    голубой/оранжевый/серый, метки у рядов, без шкалы значений и сетки."""
    data = df[df["Значение"].notna()].copy()
    if not len(data):
        return alt.Chart().mark_text(
            text=["Нет данных"], size=12, color=style["muted"]
        ).properties(width=500, height=60)
    years = list(dict.fromkeys(data["Год"]))
    colors = {str(y): style["series"][i % len(style["series"])]
              for i, y in enumerate(years)}
    order = list(dict.fromkeys(data["Показатель"]))
    base = alt.Chart(data).mark_bar(size=16, cornerRadiusEnd=2).encode(
        x=alt.X("Показатель:N", sort=order, axis=alt.Axis(
            labelColor=style["muted"], labelFont=_font(style), labelFontSize=10,
            labelAngle=0, labelLimit=220, title=None, domain=False)),
        y=alt.Y("Значение:Q", axis=None),
        xOffset="Год:N",
        color=alt.Color("Год:N", scale=alt.Scale(domain=list(colors.keys()),
                                                 range=list(colors.values())),
                        legend=alt.Legend(title="Год", orient="top",
                                          labelColor=style["muted"])),
        tooltip=[alt.Tooltip("Показатель:N", title="Показатель"),
                 alt.Tooltip("Год:N"),
                 alt.Tooltip("Значение:Q", title=value_title, format=",.2f")])
    # Метки у рядов: над столбцом для положительных, под столбцом для отрицательных
    pos = data[data["Значение"] >= 0]
    neg = data[data["Значение"] < 0]
    if len(neg):
        txt = alt.layer(
            alt.Chart(pos).mark_text(fontSize=10, dy=-4,
                                     color=style["muted"]).encode(
                x=alt.X("Показатель:N", sort=order), y=alt.Y("Значение:Q"),
                xOffset="Год:N", text=alt.Text("Метка:N")),
            alt.Chart(neg).mark_text(fontSize=10, dy=11,
                                     color=style["muted"]).encode(
                x=alt.X("Показатель:N", sort=order), y=alt.Y("Значение:Q"),
                xOffset="Год:N", text=alt.Text("Метка:N")))
    else:
        txt = alt.Chart(pos).mark_text(fontSize=10, dy=-4,
                                       color=style["muted"]).encode(
            x=alt.X("Показатель:N", sort=order), y=alt.Y("Значение:Q"),
            xOffset="Год:N", text=alt.Text("Метка:N"))
    chart = base + txt
    return _finish(chart, style, title=title, height=330)