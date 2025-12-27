import re
import os
import sys
import argparse
import pandas as pd

import fitz  # PyMuPDF
from PIL import Image
import pytesseract

BN2EN = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")

def bn_to_en_digits(s: str) -> str:
    return s.translate(BN2EN)

def looks_garbled(text: str) -> bool:
    """
    Detect CID/garbage extraction (common with Bangla PDFs).
    """
    if not text:
        return True
    low = text.lower()
    if "cid:" in low:
        return True
    # if too few Bangla chars, likely garbage
    bn_chars = sum(1 for c in text if "\u0980" <= c <= "\u09ff")
    return bn_chars < 30

def ocr_page_to_text(page: fitz.Page, dpi: int = 300) -> str:
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    # Bangla OCR (ben)
    return pytesseract.image_to_string(img, lang="ben")

def extract_text(doc: fitz.Document, force_ocr: bool = False, dpi: int = 300) -> str:
    all_text = []
    for page in doc:
        if force_ocr:
            t = ocr_page_to_text(page, dpi=dpi)
        else:
            t = page.get_text("text")
            # fallback to OCR if garbled
            if looks_garbled(t):
                t = ocr_page_to_text(page, dpi=dpi)
        all_text.append(t)
    return "\n".join(all_text)

def parse_records(full_text: str):
    """
    Parse Bangla voter list blocks.
    Expected fields:
    ####. নাম:
    ভোটার নং:
    পিতা:
    মাতা:
    পেশা: ...,জন্ম তারিখ: dd/mm/yyyy
    ঠিকানা:
    """
    text = bn_to_en_digits(full_text)

    # Split by record start "0001. নাম:"
    start_re = re.compile(r"(?m)^\s*(\d{4})\.\s*নাম:\s*")
    starts = list(start_re.finditer(text))

    records = []
    for idx, m in enumerate(starts):
        serial = int(m.group(1))
        start = m.start()
        end = starts[idx + 1].start() if idx + 1 < len(starts) else len(text)
        block = text[start:end].strip()

        # Skip migrated placeholders
        if "মাইগ্রেট" in block:
            continue

        # Name is the first line after "নাম:"
        name = ""
        first_line = block.splitlines()[0]
        # after "নাম:"
        if "নাম:" in first_line:
            name = first_line.split("নাম:", 1)[1].strip()

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

    # sort by serial
    records.sort(key=lambda x: x["Serial"])
    return records

def find_field(block: str, field_name_bn: str) -> str:
    # matches "field_name_bn: value"
    m = re.search(rf"{re.escape(field_name_bn)}:\s*(.*?)(?:\n|$)", block)
    return m.group(1).strip() if m else ""

def main():
    ap = argparse.ArgumentParser(description="Bangla Voter PDF -> Excel/CSV")
    ap.add_argument("--pdf", default=None, help="PDF file path (default: first .pdf in folder)")
    ap.add_argument("--out", default="output.xlsx", help="Output Excel file name")
    ap.add_argument("--csv", default="output.csv", help="Output CSV file name")
    ap.add_argument("--force-ocr", action="store_true", help="Always use OCR")
    ap.add_argument("--dpi", type=int, default=300, help="OCR DPI (default 300)")
    args = ap.parse_args()

    pdf_path = args.pdf
    if not pdf_path:
        # pick first pdf in current folder
        pdfs = [f for f in os.listdir(".") if f.lower().endswith(".pdf")]
        if not pdfs:
            print("❌ No PDF found in this folder. Use --pdf yourfile.pdf")
            sys.exit(1)
        pdf_path = pdfs[0]

    if not os.path.exists(pdf_path):
        print(f"❌ PDF not found: {pdf_path}")
        sys.exit(1)

    doc = fitz.open(pdf_path)
    full_text = extract_text(doc, force_ocr=args.force_ocr, dpi=args.dpi)

    records = parse_records(full_text)
    if not records:
        print("❌ No records found. Try --force-ocr")
        sys.exit(1)

    df = pd.DataFrame(records)

    # Save
    df.to_excel(args.out, index=False)
    df.to_csv(args.csv, index=False, encoding="utf-8-sig")

    print(f"✅ Done! Excel: {args.out}")
    print(f"✅ Done! CSV : {args.csv}")
    print(f"✅ Total rows: {len(df)}")

if __name__ == "__main__":
    main()
