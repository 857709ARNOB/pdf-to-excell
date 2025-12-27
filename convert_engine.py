import re
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import pytesseract

BN2EN = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")

def bn_to_en_digits(s: str) -> str:
    return s.translate(BN2EN)

def looks_garbled(text: str) -> bool:
    if not text:
        return True
    low = text.lower()
    if "cid:" in low:
        return True
    bn_chars = sum(1 for c in text if "\u0980" <= c <= "\u09ff")
    return bn_chars < 30

def ocr_page_to_text(page: fitz.Page, dpi: int = 300) -> str:
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return pytesseract.image_to_string(img, lang="ben")

def extract_text_from_pdf(pdf_path: str, force_ocr: bool = False, dpi: int = 300) -> str:
    doc = fitz.open(pdf_path)
    chunks = []
    for page in doc:
        if force_ocr:
            t = ocr_page_to_text(page, dpi=dpi)
        else:
            t = page.get_text("text")
            if looks_garbled(t):
                t = ocr_page_to_text(page, dpi=dpi)
        chunks.append(t)
    return "\n".join(chunks)

def find_field(block: str, field_name_bn: str) -> str:
    m = re.search(rf"{re.escape(field_name_bn)}:\s*(.*?)(?:\n|$)", block)
    return m.group(1).strip() if m else ""

def parse_records(full_text: str):
    text = bn_to_en_digits(full_text)

    start_re = re.compile(r"(?m)^\s*(\d{4})\.\s*নাম:\s*")
    starts = list(start_re.finditer(text))

    records = []
    for idx, m in enumerate(starts):
        serial = int(m.group(1))
        start = m.start()
        end = starts[idx + 1].start() if idx + 1 < len(starts) else len(text)
        block = text[start:end].strip()

        if "মাইগ্রেট" in block:
            continue

        first_line = block.splitlines()[0] if block.splitlines() else ""
        name = first_line.split("নাম:", 1)[1].strip() if "নাম:" in first_line else ""

        voter_no = find_field(block, "ভোটার নং")
        father = find_field(block, "পিতা")
        mother = find_field(block, "মাতা")
        address = find_field(block, "ঠিকানা")

        prof, dob = "", ""
        m_prof = re.search(r"পেশা:\s*(.*?)(?:,?\s*জন্ম তারিখ:\s*([0-9/]+))?(?:\n|$)", block)
        if m_prof:
            prof = (m_prof.group(1) or "").strip()
            dob = (m_prof.group(2) or "").strip()

        records.append({
            "Serial": serial,
            "নাম": name,
            "ভোটার নং": voter_no,
            "পিতা": father,
            "মাতা": mother,
            "পেশা": prof,
            "জন্ম তারিখ": dob,
            "ঠিকানা": address
        })

    records.sort(key=lambda x: x["Serial"])
    return records

def pdf_to_excel(pdf_path: str, excel_out: str, csv_out: str | None = None, force_ocr: bool = True):
    text = extract_text_from_pdf(pdf_path, force_ocr=force_ocr, dpi=300)
    records = parse_records(text)
    if not records:
        raise RuntimeError("No records found. Try force OCR.")

    df = pd.DataFrame(records)
    df.to_excel(excel_out, index=False)

    if csv_out:
        df.to_csv(csv_out, index=False, encoding="utf-8-sig")

    return len(df)
