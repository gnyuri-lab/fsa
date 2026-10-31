import os
import shutil
import tempfile

import pandas as pd
import streamlit as st

from data_extraction import DataExtractor
from analytics import AnalyticsEngine
from charts import (BASE_STYLE, SMALL_LABEL, accent_colors, merge_style,
                    structure_chart, dynamics_bars, fin_results_bars,
                    grouped_histogram)
from report_gen import ReportGenerator, verify_anonymity, default_out_dir

# Минимум форм для расчета (решение 30.10.2026): Баланс + ОФР;
# ОДДС и ОИК необязательны. Зафиксировано в инструкции (user_manual).
MIN_REQUIRED = ("Баланс", "ОФР")
OPTIONAL = ("ОДДС", "ОИК")

# Настройка страницы
st.set_page_config(
    page_title="FSA Анализ отчетности",
    page_icon="📊",
    layout="wide"
)

# --- Пути к ресурсам интерфейса -----------------------------------------
PROJ_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
FORMS_DIR = os.path.join(PROJ_ROOT, "forms")
FORMS_SHOTS_DIR = os.path.join(FORMS_DIR, "screenshots")
SERVICES_DIR = os.path.join(PROJ_ROOT, "information services")

def user_manual():
    """Инструкция для пользователя (экраны, кнопки, ограничения)."""
    st.markdown(
        """
        **Как работает сервис (5 экранов):**

        **Экран 1. Загрузка документов**
        - Загрузите формы отчетности: PDF (с текстовым слоем) или XLSX.
        - **Минимум для расчета — две формы: Баланс (форма 1) и ОФР
          (форма 2).** Приложение не приступит к расчету без хотя бы
          одной из них.
        - ОДДС (форма 5) и ОИК (форма 4) **необязательны**: показатели,
          зависящие от них, будут отмечены «требует уточнения».
        - Кнопка «Обработать все файлы» извлекает данные по кодам строк
          форм (Баланс, ОФР, ОДДС, ОИК).
        - Повторная загрузка файла одной формы заменяет ранее извлеченные
          строки этой формы.
        - Ниже — скриншоты поддерживаемых бланков и ссылки на государственный
          ресурс БФО (ФНС) и Smart-Lab, где можно получить отчетность.

        **Экран 2. Верификация данных**
        - Проверьте извлеченные цифры: редактируйте значения прямо в
          таблицах (кнопки настройки столбцов видны постоянно).
        - Кнопка «⬇ Скачать таблицу (.xlsx)» выгружает текущую таблицу.
        - «Проверить данные» — сверка по контрольным формулам (суммы
          разделов, связка капитала и др., ОФР 2200/2300, ОДДС
          4400+4450+4490=4500).
        - **Блокируют расчет только ошибки в Балансе и ОФР.** При
          ошибках в необязательных ОДДС/ОИК исправьте значения в
          таблицах (и проверьте снова) или исключите форму кнопкой
          «🚫 Рассчитать без …» — показатели этой формы будут отмечены
          «требует уточнения» (вернуть форму можно кнопкой «Вернуть»).
        - «Подтвердить и приступить к расчету» — переход к расчетам;
          кнопка активна, только если проверка пройдена (или лишние
          формы исключены).

        **Экран 3. Данные по амортизации**
        - Введите амортизацию за отчетный и предыдущий период (для EBITDA)
          или выберите «Нет данных» — тогда EBITDA/EBIT будет рассчитан
          без амортизации, с пометкой в таблице расчетов.
        - Здесь же можно загрузить корпоративный стиль (JSON) для
          диаграмм.

        **Что такое «стиль JSON» (Экран ввода амортизации)**
        - Это текстовый файл с настройками цветов диаграмм. Загрузите
          его на экране 3 («🎨 Стиль оформления») — и все диаграммы
          (панель, Excel, Word, PPTX) будут перекрашены в корпоративные
          цвета. Пример файла:

          ```
          {
            "series": ["#2a78d6", "#eb6834", "#8a877f"],
            "pie_top": ["#a8e05f", "#2a78d6", "#eb6834"],
            "surface": "#f4f3f0",
            "dynamics": "#359c56"
          }
          ```

          Значение каждой строки:
        - `"series": ["#2a78d6", ...]` — цвета рядов данных по годам:
          первый цвет — отчетный год (2025), второй — предыдущий (2024),
          третий — 2023. Применяется к гистограммам «Ликвидность»,
          «Финансовая устойчивость», «Рентабельность»,
          «Оборачиваемость» и графикам «Финансовые результаты».
          Цвета задаются шестизначным кодом hex (например, `#2a78d6` —
          голубой). Значений должно быть от 1 до 3, по числу лет.
        - `"pie_top": ["#a8e05f", "#2a78d6", "#eb6834"]` — цвета ТРЕХ
          КРУПНЕЙШИХ секторов круговых диаграмм, по убыванию доли:
          салатовый → голубой → оранжевый. Остальные сектора и
          «Прочие <1%» окрашиваются оттенками `"gray_ramp"`. Задается
          от 1 до 3 hex-кодов (по числу акцентных секторов).
        - `"surface": "#f4f3f0"` — цвет фона диаграмм (светло-серый в
          примере). Ставится под графиками на аналитической панели.
        - `"dynamics": "#359c56"` — цвет ряда линейчатой диаграммы
          «Динамика изменений статей Баланса» (светло-зеленый в примере).
        - Дополнительно можно задать `"green_ramp"` и `"gray_ramp"` —
          наборы по 6 hex-оттенков (светлый → темный). `"gray_ramp"`
          красит сектора круговых вне топ-3 и «Прочие <1%».
        - **Окраска круговых (по умолчанию):** три крупнейшие категории
          — акцентные цвета `"pie_top"` (салатовый, голубой, оранжевый
          по убыванию доли), остальные — оттенки `"gray_ramp"`.
        - Если какой-то ключ не указан, содержит опечатку или неверный
          код цвета — для него автоматически подставляется базовый стиль
          приложения, данные при этом не изменяются.

        **Экран 4. Основные показатели и аналитическая панель**
        - Таблица основных показателей ФХД и коэффициенты по 6 разделам
          методологии со статусами (Безопасно/Серая зона/Риск).
        - У каждой таблицы — кнопка выгрузки в .xlsx.
        - Диаграммы: структуры активов/капитала, динамика статей,
          финансовые результаты, KPI по годам.
        - Выгрузка: не забывайте проверять, что верификация выполнена.
        - Горизонтальный анализ по всем формам (в млн руб. и %).

        **Экран 5. Генерация пакета отчетов**
        - «📊 Создать Excel», «📝 Создать Word», «📽 Создать PPTX»
          или «📦 Сгенерировать пакет».
        - Файлы пишутся в папку output/ и доступны для скачивания
          кнопками под блоком.
        - В Word диаграммы вставляются редактируемыми объектами из
          Excel (нужен установленный Microsoft Office; без него — рисунки).
        - Пакет анонимизирован: удалены названия компаний, ИНН, ФИО,
          должности, адреса и телефоны (BR-03).

        **Сценарий 2. Генерация пакета (сверху):**
        - Документы создаются ТОЛЬКО по нажатию кнопки: «📊 Создать Excel»,
          «📝 Создать Word», «📽 Создать PPTX» или «📦 Сгенерировать пакет».
        - Автоматическое создание файлов без вашего действия отсутствует:
          ничего не формируется при переходе между экранами.
        - Готовые файлы кладутся в папку output/. Перед генерацией закройте
          файлы этой папки, если они открыты у вас в Office — иначе
          запись заблокируется (появится сообщение с подсказкой).
        - Word строит таблицу показателей и диаграммы из свежего Excel:
          если explonation.xlsx в этот момент заблокирован, Word будет
          сформирован по прежней версия файла (об этом выводится
          предупреждение) — закройте Excel и создайте Word снова.
        - Годы во всех документах (панель, Excel, Word, PPTX) извлекаются
          из шапок загруженных форм отчетности: например, при загрузке
          отчета за 2023 г. файлы и панель показывают 2023–2021.
          Проверяйте период в сообщениях генерации.
        - PDF-сканы обрабатываются автоматически через OCR (Tesseract,
          язык русский); распознавание может содержать цифровые
          погрешности — подтверждайте строки на экране 2.
        - Работа только локально в браузере (localhost), данные не
          отправляются на внешние серверы.
        - Минимум для расчета — 2 формы: Баланс и ОФР, проверенные
          без ошибок; ОДДС и ОИК не обязательны — при ошибках в них
          расчет можно продолжить, исключив форму (показатели
          «требует уточнения»), либо исправив данные.
        - Статьи с долей менее 1% на круговых диаграммах объединяются
          в категорию «Прочие <1%» (на панели, в Excel, Word и PPTX);
          три крупнейшие категории окрашиваются акцентными цветами
          pie_top (салатовый, голубой, оранжевый), остальные —
          оттенками серого.
        - Поддерживаются стандартные формы РФ (текстовые PDF и XLSX); для
          PDF-сканов (страницы-картинки без текстового слоя) автоматически
          применяется OCR (Tesseract). Ограничение: OCR-распознавание дает
          возможные цифровые погрешности (в т.ч. на курсивном шрифте) —
          все извлеченные строки выводятся на экран 2 (FR-02): проверьте
          цифры и исправьте перед расчетом. Точный путь — загрузка XLSX.
        - Все расчетные значения в Excel определены формулами, которые
          можно проверить; исходные данные подтверждаются аналитиком.
        - Сервис формирует шаблоны документов: окончательные выводы и
          рекомендации пользователь делает на основе своего опыта.
        """
    )


# --- Дизайн в стиле визитки (information services/my card.jpg) ----------
st.markdown(
    """
    <style>
      .card-hero {
        background: #232a63; border-radius: 8px; padding: 34px 30px 30px;
        color: #ffffff; margin-bottom: 18px;
      }
      .card-hero h1 { color: #ffffff; font-size: 30px; margin: 0 0 10px; }
      .card-hero p  { color: #cfd4f2; font-size: 15px; margin: 4px 0; }
      .card-hero .chip { display: inline-block; background: #f5b80f;
        color: #232a63; font-weight: 700; border-radius: 4px;
        padding: 3px 10px; font-size: 13px; margin-top: 8px; }
      .svc-card { background: #ffffff; border-radius: 6px;
        border-left: 4px solid #232a63; padding: 18px 16px 14px;
        box-shadow: 0 1px 3px rgba(35, 42, 99, .12); height: 100%; }
      .svc-card h3 { color: #232a63; font-size: 17px; margin: 0 0 6px; }
      .svc-card p  { color: #3d3f4c; font-size: 13px; margin: 2px 0 10px; }
      .svc-card a.btn-visit { background: #f5b80f; color: #232a63;
        font-weight: 700; font-size: 13px; border-radius: 4px;
        padding: 6px 12px; text-decoration: none; display: inline-block; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="card-hero">
      <h1>Анализ бухгалтерской отчетности</h1>
      <p>Извлечение данных из форм БО, расчет показателей по методологии
      (ликвидность, рентабельность, оборачиваемость, модели банкротства)
      и генерация пакета анонимизированных документов.</p>
      <p>Внизу страницы загрузите файлы отчетности (PDF или XLSX) —
      обработка нескольких часов сводится к нескольким минутам.</p>
      <span class="chip">Ш а б л о н &nbsp;д о к у м е н т о в</span>
    </div>
    """,
    unsafe_allow_html=True,
)
st.warning(
    "**Важно:** сервис формирует только **шаблоны документов** "
    "(Excel, Word, PowerPoint). Итоговые выводы и рекомендации — "
    "ориентир; пользователь должен доработать шаблоны с учетом своего "
    "профессионального опыта и сделать окончательные заключения "
    "самостоятельно на основании анализа.")

# --- Инструкция доступна на каждом экране (сверху) -----------------------
with st.expander("📚 Инструкция для пользователя"):
    user_manual()

# Инициализация сессии
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = {"Баланс": [], "ОФР": [], "ОДДС": [], "ОИК": []}
if "current_step" not in st.session_state:
    st.session_state.current_step = "upload"
if "verify_errors" not in st.session_state:
    st.session_state.verify_errors = None  # None — проверка еще не запускалась
if "form_verify_errors" not in st.session_state:
    # {форма: [ошибки]} для необязательных ОДДС/ОИК — не блокируют расчет
    st.session_state.form_verify_errors = None
if "excluded_forms" not in st.session_state:
    # {форма: сохраненные строки} — формы, исключенные из расчета по выбору
    st.session_state.excluded_forms = {}
if "extract_report" not in st.session_state:
    st.session_state.extract_report = []
if "analytics_results" not in st.session_state:
    st.session_state.analytics_results = None
if "generated" not in st.session_state:
    # dict: файл -> {"path": путь, "violations": [нарушения анонимизации]}
    st.session_state.generated = {}

# Инициализация экстрактора
extractor = DataExtractor()

STATUS_ICONS = {
    "Безопасно": "🟢 Безопасно",
    "Серая зона": "🟡 Серая зона",
    "Риск": "🔴 Риск",
    "требует уточнения": "⚠️ Требует уточнения",
    "—": "—",
}


def to_editor_table(data_list):
    """Готовит DataFrame к отображению: столбцы, русские названия, обратный маппинг.
    Возвращает (display_df, inverse_rename_map)."""
    df = pd.DataFrame(data_list)
    meta_cols = [c for c in ("section", "indicator", "code") if c in df.columns]
    val_cols = [c for c in df.columns if c.startswith("val_")]
    # Для ОИК: столбцы капитала (не метаданные и не значения по годам)
    oik_cols = [c for c in df.columns if c not in meta_cols and not c.startswith("val_")]

    rename_map = {
        "section": "Раздел",
        "indicator": "Показатель",
        "code": "Код",
        "val_2025": "2025 г.",
        "val_2024": "2024 г.",
        "val_2023": "2023 г.",
    }
    display_cols = meta_cols + val_cols + oik_cols
    rename_map = {**{c: n for c, n in rename_map.items() if c in df.columns},
                  **{c: c for c in oik_cols}}
    inverse_map = {orig: new for new, orig in rename_map.items()}
    user_df = df[display_cols].rename(columns=rename_map)
    return user_df, inverse_map


def kpi_years(results):
    """Годы, встречающиеся среди значений KPI (для столбцов таблицы)."""
    years = set()
    for r in results["kpi"]:
        years.update(r["values"].keys())
    return sorted(years, reverse=True)


def _report_period_text() -> str:
    """Период, за который загружена отчетность (годы извлекаются из шапок
    форм); показывается при генерации, чтобы не спутать свежий файл
    со прежним."""
    res = st.session_state.get("analytics_results")
    if not res:
        return "—"
    ycols = kpi_years(res)
    # отчетный год первым
    return " – ".join(str(y) for y in ycols) if ycols else "—"


def run_calc(depreciation):
    engine = AnalyticsEngine(st.session_state.extracted_data)
    st.session_state.analytics_results = engine.compute_all(depreciation)
    st.session_state.current_step = "analytics"
    st.rerun()


def _df_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """DataFrame в байты .xlsx (для кнопки выгрузки таблицы)."""
    import io
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append([str(c) for c in df.columns])
    for _, row in df.iterrows():
        ws.append([v if v is not None else "" for v in list(row)])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _df_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """DataFrame в байты .xlsx (для кнопки выгрузки таблицы)."""
    import io
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append([str(c) for c in df.columns])
    for _, row in df.iterrows():
        ws.append([v if v is not None else "" for v in list(row)])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()



def _ensure_writable(path):
    """Файл, открытый в Microsoft Office, заблокирован — сообщаем понятно,
    а не «не удалось распознать таблицу в PDF»."""
    if not os.path.exists(path):
        return
    try:
        with open(path, "ab"):
            pass
    except OSError as e:
        raise RuntimeError(
            f"Файл «{os.path.basename(path)}» открыт в Microsoft Office "
            "(или занят другой программой) — закройте его и запустите "
            "генерацию снова") from e


def generate_report(kind):
    """Генерация одного файла пакета ('excel'|'docx'|'pptx') в output/."""
    files = {"excel": ("explonation.xlsx", "generate_excel"),
             "docx": ("rezult.docx", "generate_docx"),
             "pptx": ("presentation.pptx", "generate_pptx")}
    fname, method = files[kind]
    gen = ReportGenerator(st.session_state.analytics_results,
                          st.session_state.extracted_data,
                          merge_style(st.session_state.get("custom_style")),
                          dep=st.session_state.get("dep_values"))
    out_dir = default_out_dir()
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, fname)
    _ensure_writable(path)
    excel_stale = None
    try:
        if method == "generate_docx":
            # Диаграммы вставляются из explonation.xlsx (OLE): пересобираем
            # его первым; если он заблокирован открытым Excel — строим Word
            # по существующему файлу того же сеанса (и предупреждаем)
            excel_path = os.path.join(out_dir, "explonation.xlsx")
            if os.path.exists(excel_path):
                try:
                    _ensure_writable(excel_path)
                    gen.generate_excel(excel_path)
                except RuntimeError as e:
                    excel_stale = str(e)
            else:
                gen.generate_excel(excel_path)
            getattr(gen, method)(path, xlsx_path=excel_path)
        else:
            getattr(gen, method)(path)
    finally:
        gen.cleanup()
    st.session_state.generated[kind] = {
        "path": path,
        "violations": verify_anonymity(path),
        "mode": getattr(gen, "docx_chart_mode", None),
        "insert_error": getattr(gen, "insert_error", None),
        "stale_excel": excel_stale,
    }


def _to_float(x):
    try:
        v = float(x)
        return None if v != v else v
    except (TypeError, ValueError):
        return None


def kpi_row(results, name):
    for r in results["kpi"]:
        if r["name"] == name:
            return r
    return None


def kpi_status_text(results, name):
    r = kpi_row(results, name)
    return r["status"] if r else "—"


def _structure_items(df, style, ramp_map=None):
    """Готовит элементы круговой: столбцы Статья/Значение/Цвет.
    ramp_map устарел (оставлен для совместимости вызовов): окраска по
    правилу 30.10.2026 — топ-3 крупные категории «салатовый/голубой/
    оранжевый» (ключ стиля pie_top), остальные — оттенки gray_ramp.
    Статьи с долей <1% объединяются в «Прочие <1%»."""
    del ramp_map
    items_list, skipped = [], []
    val_col = [c for c in df.columns if str(c).startswith("Значение, млн")][0]
    for _, r in df.iterrows():
        v = _to_float(r[val_col])
        if v is None or v <= 0:
            skipped.append(str(r["Статья"]))
            continue
        items_list.append({"Статья": str(r["Статья"]), "Значение": v,
                           "Код": str(r["Код"])})
    merged_names = []
    if items_list:
        total = sum(x["Значение"] for x in items_list)
        small_idx = [i for i, x in enumerate(items_list)
                     if total and x["Значение"] / total < 0.01]
        if small_idx:
            small = [items_list[i] for i in small_idx]
            merged_names = [x["Статья"] for x in small]
            items_list = [x for i, x in enumerate(items_list)
                          if i not in small_idx]
            items_list.append({"Статья": SMALL_LABEL,
                               "Значение": sum(x["Значение"] for x in small),
                               "Код": ""})
    if not items_list:
        return (pd.DataFrame(columns=["Статья", "Значение", "Цвет"]),
                skipped, merged_names)
    items = pd.DataFrame(items_list).sort_values("Значение", ascending=False)
    # Одинаковые названия статей в пассиве (напр. «Оценочные обязательства»
    # стр. 1430 и 1540) различаем кодом строки — иначе vega слепит их в одну
    dup = items["Статья"].value_counts()
    for name in dup[dup > 1].index:
        mask = items["Статья"] == name
        items.loc[mask, "Статья"] = name + " (стр. " + items.loc[mask, "Код"] + ")"
    # Топ-3 — акцентные цвета, остальные — серые оттенки (30.10.2026)
    items["Цвет"] = accent_colors(
        list(zip(items["Статья"], items["Значение"])), style)
    return items[["Статья", "Значение", "Цвет"]], skipped, merged_names


def _kpi_long(results, names_map, fmt):
    rows = []
    for r in results["kpi"]:
        if r["name"] in names_map:
            for y, v in r["values"].items():
                if v is not None:
                    rows.append({"Показатель": names_map[r["name"]],
                                 "Год": str(y), "Значение": v, "Метка": fmt(v)})
    return pd.DataFrame(rows)


def _dynamics_items(df):
    rows = []
    for _, r in df.iterrows():
        ch = _to_float(r["Изменение, млн руб."])
        if ch is None:
            continue
        pct = str(r.get("Изменение, %", ""))
        label = f"{ch:+,.1f} млн ({pct}%)" if pct and pct != "требует уточнения" \
            else f"{ch:+,.1f} млн"
        rows.append({"Код": r["Код"], "Статья": str(r["Статья"]),
                     "Изм": ch, "Метка": label.replace(",", " ")})
    return pd.DataFrame(rows)


def render_dashboard(results, style):
    ycols = kpi_years(results)
    year = ycols[0] if ycols else None
    if year is None:
        return
    year_col = f"{year} г., млн руб."

    # --- Карточки ключевых показателей ---
    mt = results["main_table"]

    def mt_value(name):
        sel = mt[mt["Показатель"] == name]
        return sel[year_col].iloc[0] if not sel.empty else "требует уточнения"

    def mt_fmt(name):
        v = mt_value(name)
        return v if isinstance(v, str) else f"{v:,.1f}".replace(",", " ")

    liq = kpi_row(results, "Коэффициент текущей ликвидности")
    auton = kpi_row(results, "Коэффициент автономии")
    z5 = kpi_row(results, "Z-счет Альтмана (пятифакторная модель)")
    z5_val = None
    if z5 is not None:
        for y in ycols:
            if z5["values"].get(y) is not None:
                z5_val = round(z5["values"][y], 2)
                break
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Выручка, млн руб.", mt_fmt("Выручка"))
    with c2:
        st.metric("Чистая прибыль (убыток), млн руб.", mt_fmt("Чистая прибыль (убыток)"))
    with c3:
        liq_val = None if liq is None else liq["values"].get(year)
        st.metric("Текущая ликвидность",
                  "требует уточнения" if liq_val is None else f"{liq_val:.2f}".replace(".", ","),
                  help=f"Рекомендуемый диапазон: {liq['range']}. Статус: {kpi_status_text(results, liq['name'])}"
                       if liq else "—")
    c4, c5 = st.columns(2)
    with c4:
        auton_val = None if auton is None else auton["values"].get(year)
        st.metric("Коэффициент автономии",
                  "требует уточнения" if auton_val is None else f"{auton_val:.2f}".replace(".", ","),
                  help=f"Норматив: >= 0,5. Статус: {kpi_status_text(results, 'Коэффициент автономии')}")
    with c5:
        st.metric("Z-счет Альтмана (5-факторная)",
                  "требует уточнения" if z5_val is None else f"{z5_val:.2f}".replace(".", ","),
                  help=(z5["status"] if z5 else "—") + ". "
                       "Зоны: > 2,9 — зеленая, 1,8–2,9 — серая, < 1,8 — красная")

    # --- Диаграммы по брифу ---
    cd = results["chart_data"]
    has_notes = []

    def chip(label, value, color):
        st.markdown(
            f'<div style="background:{color};border-radius:4px;padding:10px 6px;'
            f'text-align:center;font-size:13px;color:#1a1a19;">{label}<br>'
            f'<b>{value}</b></div>', unsafe_allow_html=True)

    # Плашки итогов разделов (как в шаблоне заключения)
    def mt_row(label):
        sel = mt[mt["Показатель"] == label]
        if sel.empty or isinstance(sel[year_col].iloc[0], str):
            return None
        return float(sel[year_col].iloc[0])

    def fmt_int(x):
        return "требует уточнения" if x is None else f"{x:,.0f}".replace(",", " ")

    assets_df = cd.get("Структура активов")
    if assets_df is not None:
        items, skipped, merged = _structure_items(
            assets_df, style, {"11": "gray_ramp", "12": "green_ramp"})
        if skipped:
            has_notes.append(f"«Структура активов»: статьи с нулевым/отрицательным значением "
                             f"не отображаются на круговой: {', '.join(skipped)}")
        if merged:
            has_notes.append(f"«Структура активов»: статьи с долей <1% объединены "
                             f"в «{SMALL_LABEL}»: {', '.join(merged)}")
        if not items.empty:
            c1, c2, c3 = st.columns([1, 2, 1])
            with c1:
                chip("Внеоборотные, млн руб.", fmt_int(mt_row("Внеоборотные активы")),
                     "#b3b1a9")
            with c3:
                chip("Оборотные, млн руб.", fmt_int(mt_row("Оборотные активы")),
                     style["green_ramp"][0])
            st.altair_chart(structure_chart("Структура активов, млн руб.", items, style),
                            use_container_width=False)

    cap_df = cd.get("Структура капитала")
    if cap_df is not None:
        items, skipped, merged = _structure_items(cap_df, style, {"13": "green_ramp"})
        if skipped:
            has_notes.append(f"«Структура капитала»: статьи с нулевым/отрицательным значением "
                             f"не отображаются на круговой: {', '.join(skipped)}")
        if merged:
            has_notes.append(f"«Структура капитала»: статьи с долей <1% объединены "
                             f"в «{SMALL_LABEL}»: {', '.join(merged)}")
        if not items.empty:
            c1, c2, c3 = st.columns([1, 2, 1])
            with c1:
                chip("Собственный, млн руб.", fmt_int(mt_row("Собственный капитал")),
                     style["green_ramp"][0])
            with c3:
                z = mt_row("Долгосрочные обязательства")
                k = mt_row("Краткосрочные обязательства")
                chip("Заемный, млн руб.", fmt_int(None if z is None or k is None else z + k),
                     "#b3b1a9")
            st.altair_chart(structure_chart("Структура капитала, млн руб.", items, style),
                            use_container_width=False)

    dyn_df = cd.get("Динамика изменений статей Баланса")
    if dyn_df is not None:
        rows = _dynamics_items(dyn_df)
        if not rows.empty:
            st.altair_chart(dynamics_bars(rows, style), use_container_width=True)

    fin_df = cd.get("Финансовые результаты")
    if fin_df is not None and len(ycols) > 1:
        st.altair_chart(fin_results_bars(fin_df, style, year, ycols[1]),
                        use_container_width=True)

    for label, names_map, fmt, value_title in [
        ("Ликвидность", {
            "Коэффициент обеспеченности собственными источниками": "Обеспеч. СОС\n(0,2–0,5)",
            "Коэффициент текущей ликвидности": "Текущая\n(1,5–2,0)",
            "Коэффициент быстрой ликвидности": "Быстрая\n(0,8–1,0)",
            "Коэффициент абсолютной ликвидности": "Абсолютная\n(0,3–0,5)",
        }, lambda v: f"{v:.2f}".replace(".", ","), "коэффициент"),
        ("Финансовая устойчивость", {
            "Коэффициент автономии": "Независимость\n(не менее 0,5)",
            "Коэффициент финансовой устойчивости": "Устойчивость\n(0,6–0,8)",
            "Коэффициент финансового рычага": "Рычаг\n(не более 1,0)",
        }, lambda v: f"{v:.2f}".replace(".", ","), "коэффициент"),
        ("Рентабельность", {
            "Рентабельность продаж (ROS), %": "ROS",
            "Рентабельность активов (ROA), %": "ROA",
            "Рентабельность собственного капитала (ROE), %": "ROE",
            "Рентабельность по EBITDA, %": "EBITDA",
        }, lambda v: f"{v:.1f}%".replace(".", ","), "%"),
        ("Оборачиваемость, дней", {
            "Оборачиваемость дебиторской задолженности, дней": "ДЗ",
            "Оборачиваемость кредиторской задолженности, дней": "КЗ",
            "Оборачиваемость запасов, дней": "Запасы",
            "Оборачиваемость оборотного капитала, дней": "Оборотн. кап.",
            "Цикл конверсии денежных средств, дней": "ЦКДС",
        }, lambda v: f"{v:.0f}", "дней"),
    ]:
        long_df = _kpi_long(results, names_map, fmt)
        if not long_df.empty:
            st.altair_chart(grouped_histogram(long_df, style, label, value_title),
                            use_container_width=True)

    for note in has_notes:
        st.caption("ℹ️ " + note)


# --- ЭКРАН 1: ЗАГРУЗКА ---
if st.session_state.current_step == "upload":
    st.subheader("1. Загрузка документов")
    st.info("Загрузите файлы БО (PDF или XLSX). **Минимум для расчета — "
            "2 формы: Баланс (форма 1) и ОФР (форма 2).** ОДДС и ОИК "
            "необязательны: показатели, зависящие от них, будут отмечены "
            "«требует уточнения». Повторная загрузка данных формы "
            "заменяет ранее извлеченные строки этой формы.")

    uploaded_files = st.file_uploader(
        "Выберите файлы", type=["pdf", "xlsx"], accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("🚀 Обработать все файлы"):
            report = []
            tmpdir = tempfile.mkdtemp(prefix="fsa_")
            try:
                with st.spinner("Обработка документов, пожалуйста, подождите..."):
                    for uploaded_file in uploaded_files:
                        path = os.path.join(tmpdir, uploaded_file.name)
                        with open(path, "wb") as f:
                            f.write(uploaded_file.getbuffer())

                        data = extractor.get_structured_data(path)
                        if not data:
                            report.append(f"«{uploaded_file.name}»: не удалось извлечь данные")
                            continue
                        for form_name, rows in data.items():
                            if not rows:
                                continue
                            # Данные формы перезаписываются, а не накапливаются —
                            # чтобы строки от разных загрузок не смешивались.
                            st.session_state.extracted_data[form_name] = rows
                            report.append(
                                f"«{uploaded_file.name}» → {form_name}: извлечено {len(rows)} строк"
                            )
            finally:
                shutil.rmtree(tmpdir, ignore_errors=True)

            st.session_state.extract_report = report
            st.session_state.verify_errors = None
            st.session_state.form_verify_errors = None
            st.session_state.excluded_forms = {}
            st.session_state.current_step = "verify"
            st.rerun()

    # --- Формы отчетности (скриншоты бланков, папка forms/) --- #
    st.markdown("---")
    st.subheader("Формы отчетности, поддерживаемые сервисом")
    st.caption("Официальные бланки (ФСБУ 4/2023): приложение извлекает "
               "значения по кодам строк. Наведите курсор на таблицы на "
               "следующих экранах — кнопки редактирования/сортировки "
               "столбцов видны у каждой таблицы, а таблицу можно выгрузить "
               "в .xlsx.")
    st.markdown(
        """
        <style>
          .form-card { background:#ffffff; border-radius:6px;
            border-left:4px solid #232a63; padding:12px 14px;
            box-shadow:0 1px 3px rgba(35,42,99,.12); }
          .form-card h4 { color:#232a63; margin:0 0 4px; font-size:15px; }
          .form-card p { color:#3d3f4c; font-size:12px; margin:0; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    form_cards = [
        ("Баланс_1", "Бухгалтерский баланс",
         "ОКУД 0710001: актив и пассив, коды строк 1110–1700"),
        ("ОФР_2", "Отчет о финансовых результатах",
         "ОКУД 0710002: выручка, расходы, прибыль, коды 2110–2410"),
        ("ОДДС_5", "Отчет о движении денежных средств",
         "ОКУД 0710005: потоки от операций, коды 4101–4500"),
        ("ОИК_4", "Отчет об изменении капитала",
         "ОКУД 0710004: движение капитала, коды 3100–3399"),
    ]
    cols = st.columns(4)
    for col, (png, title, subtitle) in zip(cols, form_cards):
        with col:
            img_path = os.path.join(FORMS_SHOTS_DIR, f"{png}.png")
            if os.path.exists(img_path):
                st.image(img_path, use_container_width=True)
            st.markdown(
                f'<div class="form-card"><h4>{title}</h4><p>{subtitle}</p></div>',
                unsafe_allow_html=True)

    # --- Сервисы получения отчетности ----------------------------------- #
    st.markdown("---")
    st.subheader("Сервисы получения отчетности")
    st.caption("Нажмите кнопку — откроется сайт сервиса в новой вкладке.")
    fnns_shot = os.path.join(SERVICES_DIR, "Ресурс БФО — ФНС.jpg")
    smart_shot = os.path.join(SERVICES_DIR,
                              "Фундаментальный анализ  — Smart-Lab .jpg")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            '<div class="svc-card"><h3>Государственный ИР БФО (ФНС)</h3>'
            '<p>Официальная бухгалтерская (финансовая) отчетность '
            'организаций: поиск по ИНН/названию, выгрузка форм.</p>'
            '<a class="btn-visit" href="https://bo.nalog.gov.ru/" '
            'target="_blank" rel="noopener">Открыть сервис ФНС →</a>'
            "</div>",
            unsafe_allow_html=True)
        if os.path.exists(fnns_shot):
            st.image(fnns_shot, use_container_width=True,
                     caption="Ресурс БФО — ФНС (bo.nalog.gov.ru)")
        st.link_button("🏛 Сервис ФНС (БФО)",
                       "https://bo.nalog.gov.ru/",
                       use_container_width=True)
    with col2:
        st.markdown(
            '<div class="svc-card"><h3>Smart-Lab — фундаментальный анализ</h3>'
            '<p>Показатели и мультипликаторы акций компаний РФ: '
            'выручка, прибыль, P/E, P/S, EV/EBITDA и др.</p>'
            '<a class="btn-visit" '
            'href="https://smart-lab.ru/q/shares_fundamental4/" '
            'target="_blank" rel="noopener">Открыть Smart-Lab →</a>'
            "</div>",
            unsafe_allow_html=True)
        if os.path.exists(smart_shot):
            st.image(smart_shot, use_container_width=True,
                     caption="Фундаментальный анализ — Smart-Lab")
        st.link_button("📈 Сервис Smart-Lab",
                       "https://smart-lab.ru/q/shares_fundamental4/",
                       use_container_width=True)

# --- ЭКРАН 2: ВЕРИФИКАЦИЯ ---
elif st.session_state.current_step == "verify":
    st.subheader("2. Верификация данных")
    st.write("Проверьте корректность извлеченных цифр. Вы можете редактировать значения прямо в таблице.")

    if st.session_state.extract_report:
        st.caption("Результат извлечения:")
        for line in st.session_state.extract_report:
            st.caption(f"— {line}")

    if any(st.session_state.extracted_data.values()):
        tabs = st.tabs(["Баланс", "ОФР", "ОДДС", "ОИК"])

        for i, (form_name, data_list) in enumerate(st.session_state.extracted_data.items()):
            with tabs[i]:
                if not data_list:
                    st.warning(f"Данные для формы {form_name} не найдены в документе.")
                    continue

                user_df, inverse_map = to_editor_table(data_list)
                edited_df = st.data_editor(
                    user_df,
                    key=f"editor_{form_name}",
                    num_rows="dynamic",
                    use_container_width=True
                )
                # Возвращаем данные в исходный формат словарей для session_state
                back_df = edited_df.rename(columns=inverse_map)
                st.session_state.extracted_data[form_name] = back_df.to_dict("records")
                st.download_button(
                    f"⬇️ Скачать {form_name} (xlsx)",
                    _df_xlsx_bytes(edited_df), file_name=f"{form_name}.xlsx")
    else:
        st.warning("Данные не извлечены. Вернитесь к загрузке файлов.")

    # ОДДС и ОИК не обязательны (решение 30.10.2026) — предупреждение,
    # а не блокировка расчета.
    absent_optional = [f for f in OPTIONAL
                       if not (st.session_state.extracted_data.get(f) or [])]
    if absent_optional:
        st.warning(f"Форм(а)ы {', '.join(absent_optional)} не загружены — "
                   "это допустимо: расчет возможен без них (минимум — "
                   "Баланс и ОФР). Показатели, зависящие от этих форм, "
                   "будут отмечены «требует уточнения».")

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("⬅️ Вернуться к загрузке"):
            st.session_state.verify_errors = None
            st.session_state.form_verify_errors = None
            st.session_state.excluded_forms = {}
            st.session_state.current_step = "upload"
            st.rerun()
    with col2:
        if st.button("🔍 Проверить данные"):
            with st.spinner("Выполняю сверку по правилам FSA..."):
                blocker, ferrs = extractor.verify_data_integrity(
                    st.session_state.extracted_data)
            st.session_state.verify_errors = blocker
            st.session_state.form_verify_errors = ferrs

    verify_errors = st.session_state.verify_errors
    form_errors = st.session_state.form_verify_errors or {}
    # Расчет возможен: Баланс/ОФР без ошибок, и каждая из ОДДС/ОИК либо
    # прошла проверку, либо исключена кнопкой «Рассчитать без …»
    passed = (verify_errors is not None and not verify_errors
              and not form_errors)

    # Ошибки необязательных форм: сообщаем, предлагаем исправить
    # или исключить форму из расчета (решение 30.10.2026)
    if form_errors:
        for form, ferrs in form_errors.items():
            with st.expander(
                    f"⚠️ {form}: ошибки проверки — исправьте значения в "
                    f"таблицах выше или рассчитайте без «{form}»",
                    expanded=True):
                for err in ferrs:
                    st.write(f"— {err}")
                if st.button(f"🚫 Рассчитать без «{form}» "
                             f"(показатели формы — требует уточнения)"):
                    st.session_state.excluded_forms[form] = \
                        list(st.session_state.extracted_data.get(form) or [])
                    st.session_state.extracted_data[form] = []
                    with st.spinner("Выполняю переверку..."):
                        blocker, ferrs = extractor.verify_data_integrity(
                            st.session_state.extracted_data)
                    st.session_state.verify_errors = blocker
                    st.session_state.form_verify_errors = ferrs
                    st.rerun()

    # Исключенные формы: напоминание + возможность вернуть
    for form, kept_rows in st.session_state.excluded_forms.items():
        if not kept_rows:
            continue
        st.info(f"Форма «{form}» исключена из расчета: показатели по ней "
                "будут «требует уточнения»."
                + (f" Извлеченных строк: {len(kept_rows)}." if kept_rows else ""))
        if st.button(f"↩️ Вернуть «{form}» в расчет"):
            st.session_state.extracted_data[form] = kept_rows
            st.session_state.excluded_forms.pop(form, None)
            with st.spinner("Выполняю переверку..."):
                blocker, ferrs = extractor.verify_data_integrity(
                    st.session_state.extracted_data)
            st.session_state.verify_errors = blocker
            st.session_state.form_verify_errors = ferrs
            st.rerun()

    if verify_errors is None:
        st.info("Нажмите «Проверить данные» для сверки по контрольным формулам.")
    elif verify_errors:
        st.error("❌ Обнаружены ошибки:")
        for err in verify_errors:
            st.write(f"— {err}")
        st.write("Варианты исправления: исправьте значения в таблицах выше "
                 "или вернитесь к загрузке и загрузите корректный файл.")
    elif not form_errors:
        st.success("✅ Проверка завершена успешно! Можно приступить к расчету.")
    else:
        st.warning("⚠️ Обязательные формы (Баланс, ОФР) проверены! "
                   "Необязательные формы с ошибками — см. выше: исправьте "
                   "или рассчитайте без них.")

    with col3:
        if st.button("✅ Подтвердить и приступить к расчету",
                     disabled=(verify_errors is None or bool(verify_errors)
                               or bool(form_errors))):
            st.session_state.current_step = "calc_input"
            st.rerun()

# --- ЭКРАН 3: ВВОД АМОРТИЗАЦИИ (для EBITDA) ---
elif st.session_state.current_step == "calc_input":
    st.subheader("3. Данные по амортизации (для расчета EBITDA)")
    st.write("Для EBITDA нужна амортизация за отчетный и предыдущий период. "
             "Выберите один из вариантов:")

    years = AnalyticsEngine(st.session_state.extracted_data).years

    # --- FR-11: выбор корпоративного стиля ---
    with st.expander("🎨 Стиль оформления (по умолчанию — базовый стиль приложения)"):
        st.caption("Можно загрузить JSON-файл с корпоративными цветами. Ключи: "
                   "series (до 3 hex — годы: первый, второй, третий), "
                   "pie_top (до 3 hex — три крупнейших сектора круговых: "
                   "салатовый, голубой, оранжевый), "
                   "green_ramp / gray_ramp (6 hex — оттенки секторов вне "
                   "топ-3 и «Прочие <1%»), dynamics, surface. "
                   "Неуказанные/невалидные ключи заменяются базовыми.")
        st.code('{\n'
                '  "series": ["#2a78d6", "#eb6834", "#8a877f"],\n'
                '  "pie_top": ["#a8e05f", "#2a78d6", "#eb6834"],\n'
                '  "surface": "#f4f3f0",\n'
                '  "dynamics": "#359c56"\n'
                '}', language="json")
        style_file = st.file_uploader("JSON-файл с корпоративным стилем", type=["json"])
        if style_file is not None:
            try:
                import json as _json
                parsed = _json.loads(style_file.read())
                if not isinstance(parsed, dict):
                    st.warning("Файл не содержит JSON-объект — будет использован базовый стиль.")
                else:
                    st.session_state.custom_style = parsed
                    st.success("Пользовательский стиль применен (при расчете диаграмм).")
            except Exception:
                st.warning("Не удалось прочитать JSON — будет использован базовый стиль.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Вариант 1: ввести данные**")
        a = b = 0.0
        if len(years) >= 1:
            a = st.number_input(f"Амортизация за {years[0]} г., тыс. руб.",
                                min_value=0.0, value=0.0, format="%.0f")
        if len(years) >= 2:
            b = st.number_input(f"Амортизация за {years[1]} г., тыс. руб.",
                                min_value=0.0, value=0.0, format="%.0f")
        if st.button("📉 Ввести данные"):
            dep = {}
            if len(years) >= 1:
                dep[years[0]] = a
            if len(years) >= 2:
                dep[years[1]] = b
            st.session_state.dep_values = dep
            with st.spinner("Выполняю расчет показателей..."):
                run_calc(dep)

    with col2:
        st.markdown("**Вариант 2: нет данных**")
        st.write("EBITDA будет рассчитан без амортизации, "
                 "с примечанием «рассчитано без амортизации» в таблице расчетов.")
        if st.button("🚫 Нет данных"):
            st.session_state.dep_values = None
            with st.spinner("Выполняю расчет показателей..."):
                run_calc(None)

    st.markdown("---")
    if st.button("⬅️ Назад к верификации"):
        st.session_state.current_step = "verify"
        st.rerun()

# --- ЭКРАН 4: РАСЧЕТЫ (аналитическая панель; графики — Этап 3) ---
elif st.session_state.current_step == "analytics":
    results = st.session_state.analytics_results
    if not results:
        st.error("Расчеты не выполнены. Вернитесь назад.")
    else:
        st.subheader("4. Основные показатели финансово-хозяйственной деятельности")
        st.caption("Все значения в млн руб. (делением на 1000), чтобы можно было "
                   "проверить корректность расчета.")
        st.dataframe(results["main_table"], use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Скачать основные показатели (xlsx)",
            _df_xlsx_bytes(results["main_table"]),
            file_name="Основные_показатели_ФХД.xlsx")

        # --- KPI по разделам методологии ---
        st.subheader("Коэффициенты и показатели по методологии")
        st.caption("Статус рассчитан по значению отчетного года в сравнении "
                   "с рекомендуемым диапазоном (BR-01).")
        kpi = results["kpi"]
        ycols = kpi_years(results)
        sections = []
        for r in kpi:
            if r["section"] not in sections:
                sections.append(r["section"])
        for si, section in enumerate(sections):
            with st.expander(section, expanded=(si == 0)):
                rows = [r for r in kpi if r["section"] == section]
                df = pd.DataFrame({
                    "Показатель": [r["name"] for r in rows],
                    "Формула": [r["formula"] for r in rows],
                    **{f"{y} г.": [
                        "требует уточнения" if r["values"].get(y) is None
                        else r["values"][y] for r in rows]
                        for y in ycols},
                    "Диапазон (рекоменд.)": [r["range"] for r in rows],
                    "Статус": [STATUS_ICONS.get(r["status"], r["status"])
                               for r in rows],
                })
                st.dataframe(df, use_container_width=True, hide_index=True)
                st.download_button(
                    f"⬇️ Скачать {section[:40]} (xlsx)",
                    _df_xlsx_bytes(df),
                    file_name=f"KPI_{section[:40].replace(' ', '_')}.xlsx")

        # --- Горизонтальный анализ ---
        st.subheader("Горизонтальный анализ")
        horiz = results["horizontal"]
        if horiz:
            tabs = st.tabs(list(horiz.keys()))
            for i, (form, hform_df) in enumerate(horiz.items()):
                with tabs[i]:
                    st.dataframe(hform_df, use_container_width=True, hide_index=True)
                    st.download_button(
                        f"⬇️ Скачать {form} (xlsx)",
                        _df_xlsx_bytes(hform_df),
                        file_name=f"Горизонтальный_{form}.xlsx")
        else:
            st.write("Данных достаточно только за один период — динамика не рассчитывается.")

        # --- Визуализация (карточки + диаграммы по брифу) ---
        st.subheader("Визуализация")
        render_dashboard(results, merge_style(st.session_state.get("custom_style")))

        # --- Данные для диаграмм ---
        with st.expander("Данные для диаграмм (млн руб., для проверки корректности расчета)"):
            for name, cdf in results["chart_data"].items():
                st.markdown(f"**{name}**")
                st.dataframe(cdf, use_container_width=True, hide_index=True)
                c1, c2 = st.columns([3, 1])
                with c2:
                    st.download_button(
                        f"⬇️ Скачать «{name[:36]}» (xlsx)",
                        _df_xlsx_bytes(cdf),
                        file_name=f"Данные_{name[:40].replace(' ', '_')}.xlsx")

        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("⬅️ К вводу амортизации"):
                st.session_state.current_step = "calc_input"
                st.rerun()
        with col2:
            if st.button("↩️ К верификации"):
                st.session_state.current_step = "verify"
                st.rerun()

        # --- Экран 5: Генерация отчетов (Этап 4) ---
        st.subheader("5. Генерация пакета отчетов")
        st.caption("Файлы формируются в папку output/ и доступны для скачивания. "
                   "Все выходные данные анонимизированы (BR-03).")

        colx, coly, colz, colw = st.columns(4)
        kind_map = [("excel", "📊 Создать Excel"),
                    ("docx", "📝 Создать Word"),
                    ("pptx", "📽 Создать PPTX")]
        for col, (kind, label) in zip((colx, coly, colz), kind_map):
            with col:
                if st.button(label):
                    try:
                        with st.spinner("Обработка документа, пожалуйста, "
                                        "подождите..."):
                            generate_report(kind)
                    except Exception as e:
                        st.error(f"{label}: ошибка генерации — {e}")
        with colw:
            if st.button("📦 Сгенерировать пакет"):
                failures = []
                for kind, _ in kind_map:
                    try:
                        with st.spinner("Обработка документа, пожалуйста, "
                                        "подождите..."):
                            generate_report(kind)
                    except Exception as e:
                        # Отдельный файл не прерывает пакет: остальные
                        # генерируются дальше
                        pkg_names = {"excel": "explonation.xlsx",
                                     "docx": "rezult.docx",
                                     "pptx": "presentation.pptx"}
                        failures.append(
                            f"«{pkg_names.get(kind, kind)}» — {e}")
                if failures:
                    st.warning("Часть пакета не сформирована (закройте "
                               "соответствующие файлы в Office и повторите): "
                               + "; ".join(failures))
                else:
                    st.success("Пакет сформирован полностью.")

        # Результат генерации: статус анонимизации + скачивание
        names = {"excel": "explonation.xlsx", "docx": "rezult.docx",
                 "pptx": "presentation.pptx"}
        for kind, (c1, c2) in (("excel", (0, 1)), ("docx", (1, 1)),
                               ("pptx", (2, 1))):
            info = st.session_state.generated.get(kind)
            if not info:
                continue
            st.success(f"«{names[kind]}» сформирован (отчетность за "
                       f"{_report_period_text()}).")
            if kind == "docx" and info.get("stale_excel"):
                st.warning(
                    "Word построен по прежнему explonation.xlsx (он "
                    "заблокирован открытым в Office файлом и не "
                    "перегенерирован). Для свежего Excel с диаграммами: "
                    "закройте explonation.xlsx и создайте Word снова. "
                    "Причина: " + info["stale_excel"])
            if kind == "docx" and info.get("mode") == "png":
                st.warning(
                    "Диаграммы в Word вставлены как рисунки: "
                    f"Excel/Word COM недоступен"
                    + (f" ({info.get('insert_error')})"
                       if info.get("insert_error") else "")
                    + ". Для редактируемых диаграмм установите Microsoft "
                      "Excel и Word.")
            if info["violations"]:
                st.warning(f"Контроль анонимизации ({names[kind]}): найдены "
                           "подозрительные строки:"
                           + "\n".join(info["violations"]))
            else:
                st.caption(f"✔ Контроль анонимизации «{names[kind]}»: "
                           "персон. данные не найдены.")
            try:
                with open(info["path"], "rb") as fh:
                    st.download_button(f"⬇️ Скачать {names[kind]}",
                                       fh.read(), file_name=names[kind])
            except OSError:
                st.error(f"Файл «{names[kind]}» недоступен для скачивания.")

else:
    st.error("Произошла ошибка навигации.")