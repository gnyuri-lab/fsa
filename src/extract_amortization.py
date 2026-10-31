import pdfplumber
import pytesseract
from PIL import Image
import re
import json

# Path to Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Path to the PDF file
pdf_path = r'C:\Users\user\Desktop\Documents\ClaudeCLI\FSA\input\note.pdf'

def process_page(page_num):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num - 1]
            im = page.to_image(resolution=200).original
            text = pytesseract.image_to_string(im, lang='rus+eng')
            return text
    except Exception as e:
        return f"Error processing page {page_num}: {e}"

def find_amortization_in_text(text, page_num):
    results = []
    pattern = re.compile(r'амортизац\w+', re.IGNORECASE)
    lines = text.split('\n')
    for line_idx, line in enumerate(lines):
        if pattern.search(line):
            numbers = re.findall(r'(\d[\d\s]*[,\.]?\d*)', line)
            years = re.findall(r'\b(20\d{2})\b', line)
            section = "требует уточнения"
            if line_idx > 0:
                prev_line = lines[line_idx - 1].strip()
                if prev_line and len(prev_line) < 100:
                    section = prev_line
            results.append({
                'page': page_num,
                'text': line.strip(),
                'sum': numbers[0].strip() if numbers else 'требует уточнения',
                'year': years[0] if years else 'требует уточнения',
                'section': section
            })
    return results

if __name__ == "__main__":
    # To avoid timeout and memory issues, we process pages in smaller chunks
    # We'll process 10 pages at a time and print results immediately
    total_pages = 56
    all_findings = []

    for p in range(1, total_pages + 1):
        print(f"Processing page {p}...")
        text = process_page(p)
        findings = find_amortization_in_text(text, p)
        if findings:
            all_findings.extend(findings)
            print(f"Found {len(findings)} matches on page {p}")

    print("--- FINAL FINDINGS ---")
    print(json.dumps(all_findings, ensure_ascii=False, indent=2))
    print("--- END FINDINGS ---")
