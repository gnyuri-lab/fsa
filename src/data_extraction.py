import math
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber

# --- Определение формы отчетности -------------------------------------------
# ОКУД-коды форм — самый надежный маркер (для PDF и XLSX из БФО).
OKUD_CODES = {
    "0710001": "Баланс",
    "0710002": "ОФР",
    "0710004": "ОИК",
    "0710005": "ОДДС",
}

# Ключевые слова — резервный способ, только однозначные фразы (у каждой формы
# свой заголовок, поэтому пересечений нет).
FORM_KEYWORDS = {
    "Баланс": ["бухгалтерский баланс"],
    "ОФР": ["финансовых результатах"],
    "ОИК": ["изменениях капитала", "изменения капитала"],
    "ОДДС": ["движении денежных средств", "движение денежных средств", "денежные потоки"],
}

FORMS = ["Баланс", "ОФР", "ОДДС", "ОИК"]

# Формы, без которых расчет невозможен (решение 30.10.2026)
MIN_REQUIRED = ["Баланс", "ОФР"]

# Метаколонки записей (не являются значениями)
META_COLS = {"indicator", "code", "section"}

# Допуск при сверке сумм (в тыс. руб.)
TOLERANCE = 1.0


def detect_form(text: Optional[str]) -> Optional[str]:
    """Определяет форму отчетности по ОКУД-коду, затем по ключевым словам."""
    s = (text or "").lower()
    for okud, name in OKUD_CODES.items():
        if okud in s:
            return name
    for name, keywords in FORM_KEYWORDS.items():
        if any(kw in s for kw in keywords):
            return name
    return None


def clean_value(value: Any) -> float:
    """Число из ячейки: пробелы/переносы строк убираются, (x) = -x, '-' = 0.
    Знак определяется по скобкам/минусу перед числом — устойчив к опечаткам
    верстки вида «((1 638 985)»."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return 0.0
    s = str(value).strip().replace("\n", "").replace("\xa0", "").replace(" ", "")
    if s in ("", "-", "—", "–", "(-)"):
        return 0.0
    m = re.search(r"\d+(?:[.,]\d+)?", s)
    if not m:
        return 0.0
    prefix = s[:m.start()]
    val = float(m.group().replace(",", "."))
    return -val if ("(" in prefix or "-" in prefix) else val


def norm_code(value: Any) -> str:
    """Код строки: '3312', '3312.0', 3312 -> '3312'; пустое/NaN -> ''."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    s = str(value).strip().replace("\n", "").replace(" ", "")
    if s.lower() in ("", "nan", "none"):
        return ""
    m = re.match(r"^(\d{4,5})(?:\.0+)?$", s)
    return m.group(1) if m else s


def num(value: Any) -> float:
    """Безопасное приведение значения (в т.ч. NaN из редактора) к числу."""
    try:
        if value is None or (isinstance(value, float) and math.isnan(value)):
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


# --- ОИК: канонические столбцы капитала и названия разделов -------------------
OIK_CAPITAL_COLS = [
    "Уставный капитал",
    "Собственные акции, принадлежащие акционерам",
    "Накопленная дооценка ВНА",
    "Добавочный капитал",
    "Резервный капитал",
    "Нераспределенная прибыль (убыток)",
    "Итого",
]

# Порядок проверок важен: «Добавочный капитал (без накопленной дооценки)»
# содержит слово «дооценке», поэтому проверяется раньше.
_OIK_COL_PATTERNS = [
    ("устав", "Уставный капитал"),
    ("собственн", "Собственные акции, принадлежащие акционерам"),
    ("добаво", "Добавочный капитал"),
    ("дооценка", "Накопленная дооценка ВНА"),
    ("резервн", "Резервный капитал"),
    ("нераспредел", "Нераспределенная прибыль (убыток)"),
    ("непокрыт", "Нераспределенная прибыль (убыток)"),
    ("итог", "Итого"),
]


def canonical_capital_col(name: Any) -> Optional[str]:
    """Приводит заголовок столбца капитала (PDF или XLSX) к каноническому виду."""
    s = str(name or "").lower().replace("\n", " ").strip()
    if not s:
        return None
    for pattern, canon in _OIK_COL_PATTERNS:
        if pattern in s:
            return canon
    return str(name).strip()


# Разделы ОИК: сам период и статьи периода находятся в строках.
OIK_SECTION_START = "На начало отчетного периода"
OIK_SECTION_MOVE_PREV = "Движение за аналогичный период прошлого года"
OIK_SECTION_MOVE_REPORT = "Движение за отчетный период"
OIK_SECTION_END = "На конец отчетного периода"

_OIK_SECTION_PATTERNS = [
    ("начало отчетного периода", OIK_SECTION_START),
    ("конец отчетного периода", OIK_SECTION_END),
    ("аналогичный период", OIK_SECTION_MOVE_PREV),
    ("отчетный период", OIK_SECTION_MOVE_REPORT),
]


def oik_section_from_text(text: Any) -> Optional[str]:
    s = str(text or "").lower().replace("\n", " ")
    for pattern, section in _OIK_SECTION_PATTERNS:
        if pattern in s:
            return section
    return None


class DataExtractor:
    """Извлечение данных форм БО из XLSX и PDF + проверка корректности."""

    # ------------------------------------------------------------------ #
    # Извлечение
    # ------------------------------------------------------------------ #

    def get_structured_data(self, file_path: str) -> Dict[str, List[Dict[str, Any]]]:
        low = str(file_path).lower()
        if low.endswith(".xlsx"):
            return self.extract_from_xlsx(file_path)
        if low.endswith(".pdf"):
            data = self.extract_from_pdf(file_path)
            if not data:
                # Текстового слоя нет (скан) — применяется OCR (FR-01, AC-16)
                data = self.extract_from_pdf_ocr(file_path)
            return data
        return {}

    # ---------------------------- XLSX ------------------------------- #

    def extract_from_xlsx(self, file_path: str) -> Dict[str, List[Dict[str, Any]]]:
        """Все листы книги по отдельности: сервис ФНС и БФО кладут
        в один XLSX несколько форм (лист-«Сведения об организации»
        без числовых таблиц пропускается). Форма каждого листа
        определяется по его тексту; лист формы уже извлечен — данные
        этой формы перезаписываются."""
        try:
            sheets = pd.read_excel(file_path, header=None, sheet_name=None)
        except Exception as e:
            print(f"Error reading XLSX {file_path}: {e}")
            return {}

        results: Dict[str, List[Dict[str, Any]]] = {}
        for sheet_name, df in sheets.items():
            full_text = " ".join(str(v) for v in df.values.flatten())
            form = detect_form(full_text)
            if form is None:
                continue
            if form == "ОИК":
                rows = self._oik_from_xlsx(df)
            else:
                rows = self._standard_from_xlsx(df)
            if rows:
                results[form] = rows
        return results

    @staticmethod
    def _find_code_column(df: pd.DataFrame) -> Optional[Tuple[int, int]]:
        """Строка шапки и колонка «Код» (или «Код строки»)."""
        for i in range(min(15, len(df))):
            cells = [str(c).lower() if pd.notna(c) else "" for c in df.iloc[i]]
            code_col = next((j for j, c in enumerate(cells) if "код" in c), None)
            if code_col is not None:
                return i, code_col
        return None

    @staticmethod
    def _find_name_column(df: pd.DataFrame, header_i: int,
                          code_col: int) -> int:
        """Колонка с наименованиями строк: по заголовку «Наименование…» в
        строке шапки (в файлах сервиса ФНС она далеко от колонки кода),
        иначе — колонка слева от кода."""
        cells = [str(c).lower() if pd.notna(c) else "" for c in df.iloc[header_i]]
        name_col = next((j for j, c in enumerate(cells)
                         if "наименование" in c), None)
        if name_col is not None:
            return name_col
        return code_col - 1 if code_col > 0 else 0

    @staticmethod
    def _extract_year_cols(header_row: List[Any], code_col: int) -> Dict[int, str]:
        """Колонки значений по годам из шапки: 'На 31.12.2025' -> val_2025."""
        year_cols = {}
        for j in range(code_col + 1, len(header_row)):
            m = re.search(r"(20\d\d)", str(header_row[j] or ""))
            if m:
                year_cols[j] = f"val_{m.group(1)}"
        if not year_cols:  # fallback: позиционно 2025, 2024, 2023
            years = ("2025", "2024", "2023")
            year_cols = {
                j: f"val_{y}"
                for j, y in zip(range(code_col + 1, len(header_row)), years)
            }
        return year_cols

    def _standard_from_xlsx(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Баланс / ОФР / ОДДС: показатель | код | значения по годам."""
        head = self._find_code_column(df)
        if head is None:
            return []
        header_i, code_col = head
        name_col = self._find_name_column(df, header_i, code_col)

        value_cols = self._extract_year_cols(
            [df.iat[header_i, j] for j in range(df.shape[1])], code_col
        )
        if not value_cols:
            return []

        rows = []
        for i in range(header_i + 1, len(df)):
            r = df.iloc[i]
            name = str(r[name_col]).strip() if pd.notna(r[name_col]) else ""
            code = norm_code(r[code_col])
            vals = {key: clean_value(r[j]) if j < df.shape[1] else 0.0
                    for j, key in value_cols.items()}
            # Строки без кода и без значений — служебные, пропускаем;
            # у строк показателя код всегда 4–5 цифр, короткий
            # числовой код — нумерация граф шапки («1 | 2 | 3 4 5»), пропуск.
            if code and not re.fullmatch(r"\d{4,5}", code):
                continue
            if not code and not any(vals.values()):
                continue
            rows.append({"indicator": name, "code": code, **vals})
        return rows

    def _oik_from_xlsx(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """ОИК (транспонированная форма): статьи капитала — в столбцах."""
        head = self._find_code_column(df)
        if head is None:
            return []
        header_i, code_col = head

        col_map = {}
        for j in range(code_col + 1, df.shape[1]):
            canon = canonical_capital_col(df.iat[header_i, j])
            if canon:
                col_map[j] = canon
        if not col_map:
            return []

        section = OIK_SECTION_START
        rows = []
        for i in range(header_i + 1, len(df)):
            r = df.iloc[i]
            name = str(r[0]).strip() if pd.notna(r[0]) else ""
            new_section = oik_section_from_text(name)
            if new_section:
                section = new_section
                continue
            code = norm_code(r[code_col])
            vals = {canon: clean_value(r[j]) if j < df.shape[1] else 0.0
                    for j, canon in col_map.items()}
            # у строк показателя код 4–5 цифр, короткий числовой код —
            # нумерация граф шапки
            if code and not re.fullmatch(r"\d{4,5}", code):
                continue
            if not code and not any(vals.values()):
                continue  # служебная строка/подзаголовок
            rows.append({"section": section, "indicator": name, "code": code, **vals})
        return rows

    # ----------------------------- PDF ------------------------------- #

    def extract_from_pdf(self, pdf_path: str) -> Dict[str, List[Dict[str, Any]]]:
        """PDF из БФО: постранично, форма определяется по странице и
        наследуется следующими страницами той же формы."""
        results: Dict[str, List[Dict[str, Any]]] = {}
        current_form = None
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    form = detect_form(text) or current_form
                    current_form = form
                    if not form:
                        continue
                    for table in page.extract_tables() or []:
                        if form == "ОИК":
                            rows = self._oik_from_pdf_table(table)
                        else:
                            rows = self._standard_from_pdf_table(table)
                        results.setdefault(form, []).extend(rows)
        except Exception as e:
            print(f"Error extracting from PDF {pdf_path}: {e}")
        # Формы без распознанных строк не отдаем, чтобы не перезаписывать
        # ранее загруженные данные пустым результатом.
        return {form: rows for form, rows in results.items() if rows}

    @staticmethod
    def _is_numrow_row(row: List[Any]) -> bool:
        """Служебная строка нумерации граф ('1 | 2 | 3 | ...')."""
        cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
        return bool(cells) and all(len(c) <= 2 and c.isdigit() for c in cells)

    def _table_header(self, table: List[List[Any]]) -> Tuple[int, int]:
        """(индекс шапки, колонка кода)."""
        for i, row in enumerate(table[:5]):
            cells = [(c or "").lower().replace("\n", " ") for c in row]
            code_col = next((j for j, c in enumerate(cells) if "код" in c), None)
            if code_col is not None:
                return i, code_col
        return -1, -1

    def _standard_from_pdf_table(self, table: List[List[Any]]) -> List[Dict[str, Any]]:
        if not table:
            return []
        header_i, code_col = self._table_header(table)
        if header_i < 0:
            return []
        header = table[header_i]
        cells = [(c or "").lower().replace("\n", " ") for c in header]
        name_col = next(
            (j for j, c in enumerate(cells) if "наименование" in c), max(code_col - 1, 0)
        )
        value_cols = self._extract_year_cols(header, code_col)
        if not value_cols:
            return []

        rows = []

        def cell(row, j):
            return row[j] if j < len(row) else None

        for row in table[header_i + 1:]:
            if self._is_numrow_row(row):
                continue
            # Название может быть разбито между колонкой показателя и
            # колонкой продолжения (между названием и кодом).
            parts = []
            for j in range(name_col, code_col):
                txt = str(cell(row, j) or "").replace("\n", " ").strip()
                if txt:
                    parts.append(txt)
            name = " ".join(parts)
            code = norm_code(cell(row, code_col))
            vals = {key: clean_value(cell(row, j)) for j, key in value_cols.items()}
            if not code and not any(vals.values()):
                continue
            rows.append({"indicator": name, "code": code, **vals})
        return rows

    def _oik_from_pdf_table(self, table: List[List[Any]]) -> List[Dict[str, Any]]:
        if not table:
            return []
        header_i, code_col = self._table_header(table)
        if header_i < 0:
            return []
        header = table[header_i]
        cells = [(c or "").lower().replace("\n", " ") for c in header]
        name_col = next(
            (j for j, c in enumerate(cells) if "наименование" in c), max(code_col - 2, 0)
        )

        col_map = {}
        for j in range(code_col + 1, len(header)):
            canon = canonical_capital_col(header[j])
            if canon:
                col_map[j] = canon
        if not col_map:
            return []

        section = OIK_SECTION_START
        rows = []

        def cell(row, j):
            return row[j] if j < len(row) else None

        for row in table[header_i + 1:]:
            joined = " ".join(str(c or "").replace("\n", " ") for c in row)
            new_section = oik_section_from_text(joined)
            if new_section:
                section = new_section
                continue
            if self._is_numrow_row(row):
                continue
            parts = []
            for j in range(name_col, code_col):
                txt = str(cell(row, j) or "").replace("\n", " ").strip()
                if txt:
                    parts.append(txt)
            name = " ".join(parts)
            code = norm_code(cell(row, code_col))
            vals = {canon: clean_value(cell(row, j)) for j, canon in col_map.items()}
            if not code and not any(vals.values()):
                continue
            rows.append({"section": section, "indicator": name, "code": code, **vals})
        return rows

    # ------------------------------------------------------------------ #
    # Проверка корректности (правила FSA_brief.docx п.3)
    # ------------------------------------------------------------------ #

    def verify_data_integrity(self, data: Dict[str, List[Dict[str, Any]]]
                              ) -> Tuple[List[str], Dict[str, List[str]]]:
        """Сверка по контрольным формулам.

        Возвращает (blocking, form_errors):
        - blocking — ошибки Баланса/ОФР, блокирующие расчет (обязательны
          две правильно заполненные формы — решение 30.10.2026);
        - form_errors — {форма: ошибки} для ОДДС/ОИК: расчету не мешают,
          пользователь исправляет значения или исключает форму
          («Рассчитать без …», показатели формы — «требует уточнения»)."""
        blocking: List[str] = []
        for form in MIN_REQUIRED:
            rows = data.get(form) or []
            if not rows:
                blocking.append(
                    f"{form}: данные не загружены — для расчета необходимо "
                    "минимум две формы (Баланс и ОФР)")
            else:
                blocking += self._verify_standard(form, rows)
        form_errors: Dict[str, List[str]] = {}
        for form in ("ОДДС", "ОИК"):
            rows = data.get(form) or []
            if not rows:
                continue
            found = self._verify_oik(rows) if form == "ОИК" \
                else self._verify_standard(form, rows)
            if found:
                form_errors[form] = found
        return blocking, form_errors

    # ----------------------------- OCR -------------------------------- #
    # Сканы (изображения без текстового слоя), FR-01/AC-16, реализовано
    # 30.10.2026: pypdfium2 растеризует страницы (600 dpi), pytesseract
    # распознает строки (lang=rus, psm 6); из строк собираются записи
    # формы: «Название показателя + код + числа в формате X XXX XXX».
    # Баланс и ОФР (обязательные) распознаются надежнее; ОИК - необязательная
    # форма: если графы не определились, она становится "требует уточнения".
    # Ограничение: курсивный мелкий шрифт дает отдельные цифровые ошибки -
    # все строки выводятся на экран верификации (FR-02) для подтверждения.

    _DASH_RE = re.compile(r"^[-\u2014\u2013()\s.,]+$")

    @staticmethod
    def _tesseract_path() -> str:
        """Путь к tesseract.exe: TESSERACT_CMD, PATH, стандартные пути Windows."""
        import shutil
        candidates = [os.environ.get("TESSERACT_CMD"), "tesseract.exe",
                      "tesseract",
                      r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                      r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]
        for cand in candidates:
            if not cand:
                continue
            found = shutil.which(cand) or (cand if os.path.isfile(cand)
                                           else None)
            if found:
                return os.path.abspath(found)
        raise RuntimeError(
            "Tesseract OCR не найден — установите Tesseract OCR "
            "или задайте путь в переменной TESSERACT_CMD")

    def extract_from_pdf_ocr(self, pdf_path: str
                             ) -> Dict[str, List[Dict[str, Any]]]:
        """OCR-извлечение 4 форм из PDF без текстового слоя (скан)."""
        try:
            import pypdfium2 as pypdfium
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = self._tesseract_path()
        except (RuntimeError, ImportError) as e:
            print(f"OCR недоступен: {e}")
            return {}
        results: Dict[str, List[Dict[str, Any]]] = {}
        current_form = None
        years: List[int] = []
        try:
            pdf_doc = pypdfium.PdfDocument(pdf_path)
            for idx in range(len(pdf_doc)):
                gray = pdf_doc[idx].render(scale=600 / 72).to_pil() \
                    .convert("L")
                lines = self._word_lines(gray, pytesseract)
                if not years:
                    years = self._ocr_years(lines)
                joined = " ".join(ln["text"] for ln in lines)
                form = detect_form(joined) or current_form
                current_form = form
                if not form:
                    continue
                rows = self._rows_from_ocr_lines(lines, years, form)
                if rows:
                    results.setdefault(form, []).extend(rows)
        except Exception as e:
            print(f"Error OCR {pdf_path}: {e}")
        return {form: rows for form, rows in results.items() if rows}

    @staticmethod
    def _ocr_years(lines: List[Dict[str, Any]]) -> List[int]:
        """Годы шапки, слева направо (порядок графики выгрузки)."""
        for ln in lines[:15]:
            years = [int(y) for y in re.findall(r"20\d\d", ln["text"])
                     if 1990 < int(y) < 2100]
            if len(years) >= 2 and years == sorted(years, reverse=True) \
                    and " г" in ln["text"]:
                return years
        return []

    def _word_lines(self, img, pytesseract) -> List[Dict[str, Any]]:
        """Строки страницы: слова OCR сгруппированы по y-центру
        (порог = 80% медианной высоты слова)."""
        d_frame = pytesseract.image_to_data(
            img, lang="rus", config="--psm 6",
            output_type=pytesseract.Output.DATAFRAME)
        recs = []
        for _, r in d_frame.iterrows():
            token = str(r["text"]).strip()
            if not token or token == "nan" or float(r["conf"]) < 30:
                continue
            h = float(r["height"])
            cy = float(r["top"]) + h / 2
            cx = float(r["left"]) + float(r["width"]) / 2
            recs.append((cy, cx, h, token))
        if not recs:
            return []
        med_h = sorted(rr[2] for rr in recs)[len(recs) // 2]
        gap = med_h * 0.8
        recs.sort(key=lambda x: x[0])
        out_lines, cur, last_y = [], [], None
        for cy, cx, h, token in recs:
            if last_y is not None and cy - last_y > gap:
                cur.sort(key=lambda x: x[1])
                y_top = min(cy - h / 2 for cy, cx, h, t in cur)
                out_lines.append({
                    "text": " ".join(t for _, _, _, t in cur),
                    "lefts": [cx for _, cx, _, _ in cur], "y": y_top})
                cur = []
            cur.append((cy, cx, h, token))
            last_y = cy
        if cur:
            cur.sort(key=lambda x: x[1])
            y_top = min(cy - h / 2 for cy, cx, h, t in cur)
            out_lines.append({
                "text": " ".join(t for _, _, _, t in cur),
                "lefts": [cx for _, cx, _, _ in cur], "y": y_top})
        return out_lines

    def _rows_from_ocr_lines(self, lines, years, form):
        """Числовые строки -> записи формы. Число «X XXX XXX» сливается из
        токенов: внутренние группы ровно 3 цифры, крайняя 1-3; прочерк -
        пустая графа (позиция сохраняется); знак минуса из скобок/минуса."""
        rows = []
        for ln in lines:
            if self._is_numrow_row([ln["text"]]):
                continue
            tokens = ln["text"].split()
            values = []       # справа налево: float | None
            pending = ""      # накопленные внутренние группы (справа)
            while tokens:
                token = tokens[-1].strip().strip("|")
                if self._DASH_RE.match(token) or token in ("-", "—", "–"):
                    tokens.pop()
                    digits = re.sub(r"\D", "", token)
                    if digits:
                        values.append(clean_value(
                            ("-" if "(" in token else "") + digits))
                    else:
                        values.append(None)
                    continue
                bare = token.replace("\xa0", "").replace(" ", "")
                if not re.fullmatch(r"[-+]?\d+(?:[.,]\d{1,2})?", bare):
                    break  # дальше слова названия
                tokens.pop()
                if "," in bare:
                    values.append(clean_value(bare + pending))
                    pending = ""
                    continue
                if len(bare) == 3:
                    pending = bare + pending
                    continue
                values.append(clean_value(bare + pending))
                pending = ""
            if pending:
                values.append(clean_value(pending))
            if len([v for v in values if v is not None]) < 2:
                continue
            values = values[::-1]        # слева направo
            code = ""
            if tokens and re.fullmatch(r"\d{4}", tokens[-1].strip()):
                code = tokens[-1].strip()
                tokens.pop()
            else:
                for w in reversed(tokens):
                    if re.fullmatch(r"\d{4}", w.strip().strip("|")):
                        code = w.strip().strip("|")
                        break
            name = " ".join(tokens).strip(" *1.,:;")
            if not name and not code:
                continue
            record = {"indicator": name, "code": code}
            for j, value in enumerate(values):
                if value is not None:
                    if form == "ОИК":
                        record[f"ocr_{j}"] = value
                    elif j < len(years):
                        record[f"val_{years[j]}"] = value
            has_value = any(k.startswith(("val_", "ocr_")) for k in record)
            if (code or name) and has_value:
                rows.append(record)
        return rows

    @staticmethod
    def _year_cols(rows: List[Dict[str, Any]]) -> List[str]:
        cols = {c for row_rows in rows for c in row_rows
                if c.startswith("val_")}
        return sorted(cols)

    @staticmethod
    def _eq_msg(form: str, year: str, rule: str, a: float,
                b: float) -> str:
        return (f"{form} ({year}): ошибка в правиле «{rule}» — "
                f"получено {a:,.0f} против {b:,.0f} (сравниваемые значения), "
                f"требуется исправить").replace(",", " ")

    def _verify_standard(self, form: str,
                         rows: List[Dict[str, Any]]) -> List[str]:
        """Правила для Баланса, ОФР и ОДДС по каждому году."""
        errors = []
        for year_col in self._year_cols(rows):
            year = year_col[len("val_"):]
            values = {norm_code(r.get("code")): num(r.get(year_col))
                      for r in rows if norm_code(r.get("code"))}
            # Код может повторяться при ручном вводе — оставляем последнее значение.
            if form == "Баланс":
                errors += self._check_balance(year, values)
            elif form == "ОФР":
                errors += self._check_ofr(year, values)
            else:
                errors += self._check_odds(year, rows, year_col)
        return errors

    def _check_balance(self, year: str, v: Dict[str, float]) -> List[str]:
        errors = []

        def total(code):
            if code not in v:
                errors.append(f"Баланс ({year}): отсутствует строка {code} — требует уточнения")
                return None
            return v[code]

        def rng_sum(lo, hi):
            return sum(val for code, val in v.items() if lo <= int(code) <= hi)

        rules = [
            ("сумма строк 1110–1190 равна 1100", lambda t: rng_sum(1110, 1190), "1100"),
            ("сумма строк 1210–1290 равна 1200", lambda t: rng_sum(1210, 1290), "1200"),
            ("сумма строк 1100 и 1200 равна 1600", lambda t: t and t.get("1100", 0) + t.get("1200", 0), "1600"),
            ("сумма строк 1310–1390 равна 1300", lambda t: rng_sum(1310, 1390), "1300"),
            ("сумма строк 1410–1490 равна 1400", lambda t: rng_sum(1410, 1490), "1400"),
            ("сумма строк 1510–1590 равна 1500", lambda t: rng_sum(1510, 1590), "1500"),
            ("сумма строк 1300, 1400 и 1500 равна 1700",
             lambda t: t and t.get("1300", 0) + t.get("1400", 0) + t.get("1500", 0), "1700"),
        ]
        totals = {}
        for code in ("1100", "1200", "1600", "1300", "1400", "1500", "1700"):
            val = total(code)
            if val is not None:
                totals[code] = val

        for msg, calc, target in rules:
            if target not in totals:
                continue  # про отсутствие строки уже сообщено
            got = calc(totals)
            if got is None:
                continue
            if abs(got - totals[target]) > TOLERANCE:
                errors.append(self._eq_msg(f"Баланс", year, msg, got, totals[target]))

        if "1600" in totals and "1700" in totals:
            if abs(totals["1600"] - totals["1700"]) > TOLERANCE:
                errors.append(self._eq_msg(
                    "Баланс", year, "строка 1600 равна строке 1700",
                    totals["1600"], totals["1700"]))
        return errors

    def _check_ofr(self, year: str, v: Dict[str, float]) -> List[str]:
        errors = []

        def total(code):
            if code not in v:
                errors.append(f"ОФР ({year}): отсутствует строка {code} — требует уточнения")
                return None
            return v[code]

        def rng_sum(lo, hi):
            return sum(val for code, val in v.items() if lo <= int(code) <= hi)

        def need(code):
            if code not in v:
                errors.append(f"ОФР ({year}): отсутствует строка {code} — требует уточнения")
            return v.get(code, 0.0)

        # Правила для 2200 и 2300 уточнены по согласованию с пользователем:
        # буквальная формулировка брифа дает ложные ошибки на корректных данных.
        # Правила для 2400 НЕТ: строка 2410 переносится из налоговой декларации
        # без вычисления, и в строках 2411, 2412 (детализация 2410), а также в
        # строках без кода арифметический смысл отсутствует — связка с 2400
        # не гарантирована (в PDF-выгрузке БФО она заведомо не выполняется).
        checks = [
            ("сумма строк 2110–2190 равна 2100", rng_sum(2110, 2190), "2100", None),
            ("строка 2100 плюс сумма строк 2210–2290 равна 2200",
             rng_sum(2210, 2290), "2200", "2100"),
            ("строка 2200 плюс сумма строк 2310–2390 равна 2300",
             rng_sum(2310, 2390), "2300", "2200"),
        ]
        for code in ("2100", "2200", "2300"):
            need(code)

        for msg, part, target, base_code in checks:
            if target not in v:
                continue  # про отсутствие строки уже сообщено
            base = v.get(base_code, 0.0) if base_code else 0.0
            if abs(base + part - v[target]) > TOLERANCE:
                errors.append(self._eq_msg("ОФР", year, msg, base + part, v[target]))
        return errors

    def _check_odds(self, year: str, rows: List[Dict[str, Any]], year_col: str) -> List[str]:
        """ОДДС: итоги проверяются по позиции — строки блока находятся между
        итоговыми строками; строки без кода внутри блока включаются
        (согласовано с пользователем)."""
        errors = []
        ordered = [(norm_code(r.get("code")), num(r.get(year_col))) for r in rows]

        def idx(code):
            for i, (c, _) in enumerate(ordered):
                if c == code:
                    return i
            return None  # про отсутствие строки сообщается через total()

        def between_sum(a, b, lo, hi):
            ia, ib = idx(a), idx(b)
            if ia is None or ib is None or ia > ib:
                return 0.0
            return sum(val for c, val in ordered[ia + 1:ib]
                       if c == "" or lo <= int(c) <= hi)

        v = {code: val for code, val in ordered if code}

        def total(code):
            if code not in v:
                errors.append(f"ОДДС ({year}): отсутствует строка {code} — требует уточнения")
                return None
            return v[code]

        checks = [
            ("сумма строк 4111–4119 равна 4110", between_sum("4110", "4120", 4111, 4119), "4110"),
            ("сумма строк 4121–4129 равна 4120", between_sum("4120", "4100", 4121, 4129), "4120"),
            ("сумма строк 4211–4219 равна 4210", between_sum("4210", "4220", 4211, 4219), "4210"),
            ("сумма строк 4221–4229 равна 4220", between_sum("4220", "4200", 4221, 4229), "4220"),
            # (опечатка в брифе исправлена по согласованию): 4310 и 4320
            ("сумма строк 4311–4319 равна 4310", between_sum("4310", "4320", 4311, 4319), "4310"),
            ("сумма строк 4321–4329 равна 4320", between_sum("4320", "4300", 4321, 4329), "4320"),
        ]
        for code in ("4110", "4120", "4100", "4210", "4220", "4200",
                     "4310", "4320", "4300", "4400", "4500"):
            total(code)

        for msg, part, target in checks:
            if target not in v:
                continue
            if abs(part - v[target]) > TOLERANCE:
                errors.append(self._eq_msg("ОДДС", year, msg, part, v[target]))

        if all(c in v for c in ("4110", "4120", "4100")):
            if abs(v["4110"] + v["4120"] - v["4100"]) > TOLERANCE:
                errors.append(self._eq_msg(
                    "ОДДС", year, "сумма строк 4110 и 4120 равна 4100",
                    v["4110"] + v["4120"], v["4100"]))
        if all(c in v for c in ("4210", "4220", "4200")):
            if abs(v["4210"] + v["4220"] - v["4200"]) > TOLERANCE:
                errors.append(self._eq_msg(
                    "ОДДС", year, "сумма строк 4210 и 4220 равна 4200",
                    v["4210"] + v["4220"], v["4200"]))
        if all(c in v for c in ("4310", "4320", "4300")):
            if abs(v["4310"] + v["4320"] - v["4300"]) > TOLERANCE:
                errors.append(self._eq_msg(
                    "ОДДС", year, "сумма строк 4310 и 4320 равна 4300",
                    v["4310"] + v["4320"], v["4300"]))
        if all(c in v for c in ("4100", "4200", "4300", "4400")):
            if abs(v["4100"] + v["4200"] + v["4300"] - v["4400"]) > TOLERANCE:
                errors.append(self._eq_msg(
                    "ОДДС", year, "сумма строк 4100, 4200 и 4300 равна 4400",
                    v["4100"] + v["4200"] + v["4300"], v["4400"]))
        # 4400 + 4450 + 4490 = 4500 (правило согласовано 30.10.2026:
        # прежняя «сумма строк с 4400 по 4490» давала ложные ошибки).
        # Отсутствующие 4450/4490 принимаются нулевыми.
        part_fx = sum(v.get(c, 0.0) for c in ("4400", "4450", "4490"))
        if "4400" in v and "4500" in v:
            if abs(part_fx - v["4500"]) > TOLERANCE:
                errors.append(self._eq_msg(
                    "ОДДС", year, "сумма строк 4400, 4450 и 4490 равна 4500",
                    part_fx, v["4500"]))
        return errors

    def _verify_oik(self, rows: List[Dict[str, Any]]) -> List[str]:
        """ОИК: капитал на отчетную дату = капитал на конец предыдущего года
        после корректировки + движения отчетного года, по каждому столбцу
        капитала. Поддерживаются обе системы кодов строк (PDF и XLSX)."""
        errors = []
        variants = [
            # Коды файла ОИК.xlsx (бриф FSA_brief.docx п.3)
            {"closing": "83300", "opening": "63201",
             "moves": ("73311", "3312", "3327", "3330", "3316", "3340")},
            # Официальные коды PDF (ОКУД 0710004)
            {"closing": "3300", "opening": "3201",
             "moves": ("3311", "3312", "3327", "3330", "3316", "3340")},
            # Файлы сервиса ФНС: строки только 3100/3200/3300,
            # 3200 — капитал на конец прошлого года (начало отчетного);
            # движения годовыми итогами 3310/3320 (строки «в том
            # числе» идут по кодам 3311+/3321+ и не содержат всей суммы)
            {"closing": "3300", "opening": "3200",
             "moves": ("3310", "3320")},
        ]
        by_code = {}
        for r in rows:
            code = norm_code(r.get("code"))
            if code:
                by_code[code] = r

        variant = None
        for cand in variants:
            if cand["closing"] in by_code and cand["opening"] in by_code:
                variant = cand
                break
        if variant is None:
            errors.append("ОИК: не найдены итоговые строки капитала "
                          "(коды 83300/63201, 3300/3201 или 3300/3200) — "
                          "проверка невозможна, требует уточнения")
            return errors

        move_codes = [c for c in variant["moves"] if c in by_code]
        capital_cols = [c for c in rows[0] if c not in META_COLS]
        if not capital_cols:
            errors.append("ОИК: в данных отсутствуют столбцы капитала — требует уточнения")
            return errors

        closing_row = by_code[variant["closing"]]
        opening_row = by_code[variant["opening"]]
        label = {
            "83300": "83300 (ОИК.xlsx)",
            "3300": "3300 (ОИК PDF)",
        }[variant["closing"]]

        missing = [c for c in variant["moves"] if c not in by_code]
        if missing:
            errors.append(f"ОИК: отсутствуют строки движений с кодами "
                          f"{', '.join(missing)} — требует уточнения")

        for col in capital_cols:
            closing_val = num(closing_row.get(col))
            opening_val = num(opening_row.get(col))
            moves_val = sum(num(by_code[c].get(col)) for c in move_codes)
            if abs(opening_val + moves_val - closing_val) > TOLERANCE:
                errors.append(
                    f"ОИК: по столбцу «{col}» итог строки {label} ({closing_val:,.0f}) "
                    f"не равен сумме строки {variant['opening']} и движений отчетного года "
                    f"({opening_val:,.0f} + {moves_val:,.0f})".replace(",", " ")
                )
        return errors