"""Распознавание амортизации из загруженных PDF (текстовый слой + OCR).

Используется на экране «3. Данные по амортизации»: приложение ищет в
загруженных PDF строки со словом «амортизация» и числами и предлагает
найденные значения для подтверждения пользователем — расчет EBITDA
выполняется только после нажатия «Ввести данные». Текстовый слой PDF
используется напрямую; для страниц-сканов выполняется OCR (Tesseract),
точность OCR ограничена — значения обязательно проверяются пользователем.
"""
import os
import re
import shutil

import pdfplumber

# Строки, содержащие слово «амортизация» в любой форме
AMORT_RE = re.compile(r"амортизац\w+", re.IGNORECASE)
# Числа с тысячными разделителями-пробелами: «3 725 150», «2853 556»
# (ошибка OCR — «2 853 556»), обычные целые; десятичная часть через , или .
NUM_RE = re.compile(r"\d+(?:[ \u00a0]\d{3})*(?:[.,]\d{1,2})?")
YEAR_RE = re.compile(r"\b(20\d{2})\b")


def _tesseract_path() -> str:
    """Путь к tesseract.exe: TESSERACT_CMD, PATH, стандартные пути Windows
    (та же логика, что в data_extraction._tesseract_path)."""
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


def _parse_number(token: str):
    """«3 725 150» → 3725150.0; нераспознаваемый токен → None."""
    cleaned = token.replace("\u00a0", "").replace(" ", "")
    cleaned = cleaned.replace(",", ".").rstrip(".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _line_candidates(text: str, page_num: int):
    """Страница текста → строки со словом «амортизация» и числами."""
    out = []
    for line in text.split("\n"):
        if not AMORT_RE.search(line):
            continue
        sums = [n for n in (_parse_number(m) for m in NUM_RE.findall(line))
                if n is not None]
        years = [int(y) for y in YEAR_RE.findall(line)]
        out.append({"page": page_num, "text": line.strip(),
                    "sums": sums, "years": years})
    return out


def _ocr_page(pdf):
    """OCR одной страницы-скана (растеризация pdfplumber + pytesseract).

    RuntimeError (Tesseract не установлен) и ImportError (нет пакета)
    пробрасываются наружу — приложение должно объяснить это пользователю,
    а не показывать «не найдено»; остальные ошибки OCR страницы
    возвращают пустой текст."""
    try:
        import pytesseract
        pytesseract.pytesseract.tesseract_cmd = _tesseract_path()
    except (RuntimeError, ImportError):
        raise
    try:
        im = pdf.to_image(resolution=200).original
        return pytesseract.image_to_string(im, lang="rus+eng")
    except Exception:
        return ""


def extract_from_pdf(pdf_path: str, max_pages: int = 60,
                     progress_cb=None):
    """Все строки со словом «амортизация» и числами в PDF.

    Страницы с текстовым слоем читаются напрямую, пустые (сканы) — через
    OCR. progress_cb(done, total) вызывается после каждой страницы.
    max_pages ограничивает время распознавания больших файлов.

    Возвращает кортеж (кандидаты, info):
    info = {"pages": обработано страниц, "ocr_pages": распознано сканов,
            "ocr_available": доступен ли Tesseract}.
    Если Tesseract не установлен, сканы пропускаются, но страницы с
    текстовым слоем обрабатываются, а info["ocr_available"] = False.
    """
    results = []
    info = {"pages": 0, "ocr_pages": 0, "ocr_available": True}
    with pdfplumber.open(pdf_path) as pdf:
        total = min(max_pages or len(pdf.pages), len(pdf.pages))
        for idx in range(total):
            page = pdf.pages[idx]
            text = page.extract_text() or ""
            if len(text.strip()) < 20:
                try:
                    text = _ocr_page(page)
                    info["ocr_pages"] += 1
                except (RuntimeError, ImportError):
                    info["ocr_available"] = False
                    text = ""
            results.extend(_line_candidates(text, idx + 1))
            info["pages"] = idx + 1
            if progress_cb:
                progress_cb(idx + 1, total)
    return results, info


def pick_depreciation(candidates, years):
    """Предложенные значения амортизации (для подтверждения пользователем).

    Приоритет — строка, начинающаяся со слова «амортизация», где два
    числа (графы двух лет, например итоговая строка таблицы «Затраты
    на производство»: «Амортизация 3 725 150 3 485 935»): первое число
    относится к отчетному году, второе — к предыдущему. Если двучисловых
    строк нет, берется лучшая строка с одним числом (отчетный год).
    Возвращает {год: значение} или None.
    """
    relevant = [c for c in candidates if c["sums"]]
    if not relevant or not years:
        return None

    def rank(c):
        # строка, где слово «амортизация» в начале — итоговая строка таблицы
        starts = 0 if AMORT_RE.match(c["text"]) else 1
        # количество чисел желательно как число лет (максимум 2 столбца)
        fit = abs(min(len(c["sums"]), 2) - min(len(years), 2))
        return (starts, fit)

    two = [c for c in relevant if len(c["sums"]) >= 2]
    if two:
        best = min(two, key=rank)
        return {years[i]: best["sums"][i]
                for i in range(min(len(years), 2))}
    one = min(relevant, key=rank)
    return {years[0]: one["sums"][0]}