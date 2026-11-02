"""Распознавание амортизации из загруженных PDF (текстовый слой + OCR).

Используется на экране «3. Данные по амортизации»: приложение ищет в
загруженных PDF строки со словом «амортизация» и числами и предлагает
найденные значения для подтверждения пользователем — расчет EBITDA
выполняется только после нажатия «Ввести данные». Текстовый слой PDF
используется напрямую; для страниц-сканов выполняется OCR (Tesseract),
точность OCR ограничена — значения обязательно проверяются пользователем.

Второй признак — строка «Амортизация» с кодом 5640 из раздела
«Затраты на производство» табличных форм Пояснений: OCR часто выносит
код и числа в соседние строки, поэтому строка-кандидат собирается
из окна ±1 строка вокруг строки с кодом. Если же страничный OCR
теряет числа этой строки, запускается прицельный OCR полосы таблицы
с кодом (только цифровой алфавит).
"""
import os
import re
import shutil

import pdfplumber

# Строки, содержащие слово «амортизация» в любой форме; допуск
# к ошибкам OCR («Амортызация», «Амортьтзационные» и т.п.)
AMORT_RE = re.compile(r"аморт\w{0,4}зац\w+", re.IGNORECASE)
# Код строки «Амортизация» в разделе «Затраты на производство»
# Пояснений (в этой форме амортизация — строка 5640)
AMORT_CODE_RE = re.compile(r"564[0Оo]")
# Заголовок раздела Пояснений, где эта таблица (устойчив к 200 dpi)
AMORT_SECTION_RE = re.compile(r"затрат\w*\s+на\s+производство",
                              re.IGNORECASE)
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


def _line_numbers(line: str):
    """Строка → (суммы, годы): числа кроме кода 5640 и значений года
    («2021 г.» не считается значением амортизации)."""
    cuts = [m.span() for m in AMORT_CODE_RE.finditer(line)]
    sums, years = [], []
    for m in NUM_RE.finditer(line):
        if any(s <= m.start() and m.end() <= e for s, e in cuts):
            continue
        n = _parse_number(m.group())
        if n is None:
            continue
        if 2000 <= n <= 2099:
            years.append(int(n))
            continue
        sums.append(n)
    return sums, years


def _line_candidates(text: str, page_num: int):
    """Страница текста → строки со словом «амортизация» и числами,
    а также строки-кандидаты по коду 5640."""
    lines = text.split("\n")
    out = []
    for line in lines:
        if not AMORT_RE.search(line):
            continue
        sums, years = _line_numbers(line)
        years.extend(int(y) for y in YEAR_RE.findall(line))
        out.append({"page": page_num, "text": line.strip(),
                    "sums": sums, "years": years})
    # Признак «код строки 5640»: если строка с кодом и слово
    # «амортизация» в соседних строках — объединяем окно ±1 строку
    for i, line in enumerate(lines):
        if not AMORT_CODE_RE.search(line):
            continue
        if AMORT_RE.search(line):
            continue  # слово на этой же строке — уже учтено выше
        window = lines[max(0, i - 1):i + 2]
        if len(window) < 2 or \
                not any(AMORT_RE.search(w) for w in window):
            continue
        merged = " | ".join(w.strip() for w in window if w.strip())
        sums, years = _line_numbers(merged)
        years.extend(int(y) for y in YEAR_RE.findall(merged))
        out.append({"page": page_num, "text": merged, "sums": sums,
                    "years": years, "code": AMORT_CODE_RE.search(line).group()})
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


def _reocr_code_rows(pdf, page_num: int):
    """Прицельный OCR числовой части строк с кодом 5640.

    Страничный OCR сложных таблиц теряет числа строки «Амортизация»
    (текст колонок встает вертикально в общий разбор), хотя сам код
    распознается. Тогда по данным image_to_data находим слово
    «амортизация» или код 5640, вырезаем полосу этой строки таблицы
    и читаем ее отдельно (psm 7, цифровой алфавит, 2× увеличение).
    Поиск слов ведется на 400 dpi — на нижних разрешениях строка
    «Амортизация»/«5640» в сложных сканах не распознается вовсе.

    Возвращает список кандидатов (может быть пустым)."""
    try:
        import pytesseract
        from PIL import Image
        pytesseract.pytesseract.tesseract_cmd = _tesseract_path()
        im = pdf.to_image(resolution=400).original
        data = pytesseract.image_to_data(
            im, lang="rus+eng", output_type=pytesseract.Output.DICT)
    except Exception:
        return []
    W, H = im.size
    # слова для выреза: слово/код одной строки дают одинаковый top —
    # снимаем дубли, округляя до 1/50 высоты изображения
    tops = set()
    strips = []
    for i, w in enumerate(data.get("text") or []):
        if not w.strip() or not (AMORT_RE.search(w)
                                 or AMORT_CODE_RE.search(w)):
            continue
        top = data["top"][i]
        y0 = max(0, int(top - 0.4 * data["height"][i]))
        if int(y0 / (H / 50)) in tops:
            continue
        tops.add(int(y0 / (H / 50)))
        y1 = min(H, int(top + 1.7 * data["height"][i]))
        strips.append((y0, y1))
    out = []
    for y0, y1 in strips:
        strip = im.crop((0, y0, W, y1))
        strip = strip.resize((W * 2, (y1 - y0) * 2), _pil_resample())
        # цифровой whitelist склеивает числа в один токен — читаем
        # полосу обычным rus+eng, разделители-пробелы сохраняются
        txt = pytesseract.image_to_string(strip, lang="rus+eng",
                                          config="--psm 7")
        if not AMORT_RE.search(txt) and not AMORT_CODE_RE.search(txt):
            continue
        sums, years = _line_numbers(txt)
        out.append({"page": page_num, "text": txt.strip(), "sums": sums,
                    "years": years, "code": "5640"})
    return out


def _pil_resample():
    """Resampled-константа PIL: LANCZOS (новые версии) или ANTIALIAS."""
    from PIL import Image
    return getattr(Image, "Resampling", Image).LANCZOS


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
            page_cands = _line_candidates(text, idx + 1)
            # если строка-кандидат по коду 5640 есть, но чисел по нему
            # нет — прицельный OCR полосы этой строки (страничный OCR
            # теряет числа таблиц); запускается и по заголовку «Затраты
            # на производство», когда сам код на нижнем разрешении не прочитался
            if not any(c.get("code") and c["sums"] for c in page_cands) \
                    and (AMORT_CODE_RE.search(text)
                         or AMORT_SECTION_RE.search(text)):
                page_cands.extend(_reocr_code_rows(page, idx + 1))
            results.extend(page_cands)
            info["pages"] = idx + 1
            if progress_cb:
                progress_cb(idx + 1, total)
    return results, info


def pick_depreciation(candidates, years):
    """Предложенные значения амортизации (для подтверждения пользователем).

    Приоритет — строка с кодом 5640 («Амортизация» из раздела «Затраты
    на производство» Пояснений), затем строка, начинающаяся со слова
    «амортизация», где два числа (графы двух лет, например итоговая
    строка таблицы «Затраты на производство»: «Амортизация 3 725 150
    3 485 935»): первое число относится к отчетному году, второе —
    к предыдущему. Если таких строк нет, берется лучшая строка
    с одним числом (отчетный год). Возвращает {год: значение} или None.
    """
    relevant = [c for c in candidates if c["sums"]]
    if not relevant or not years:
        return None

    def rank(c):
        # строка с кодом 5640 — точная ссылка на амортизацию
        code = 0 if c.get("code") else 1
        # строка, где слово «амортизация» в начале — итоговая строка таблицы
        starts = 0 if AMORT_RE.match(c["text"]) else 1
        # количество чисел желательно как число лет (максимум 2 столбца)
        fit = abs(min(len(c["sums"]), 2) - min(len(years), 2))
        return (code, starts, fit)

    two = [c for c in relevant if len(c["sums"]) >= 2]
    if two:
        best = min(two, key=rank)
        return {years[i]: best["sums"][i]
                for i in range(min(len(years), 2))}
    one = min(relevant, key=rank)
    return {years[0]: one["sums"][0]}