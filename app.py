import os
import re
import uuid
import io

from flask import Flask, request, send_file, render_template_string
import pandas as pd
import fitz  # PyMuPDF

from PIL import Image
import pytesseract

# ======================
# CONFIG
# ======================
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)

HTML = """
<!doctype html>
<html lang="bn">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>PDF → Excel (Bangla Voter List)</title>
  <style>
    body{font-family:Arial,sans-serif;background:#0b1220;color:#e8eefc;margin:0}
    .wrap{max-width:820px;margin:40px auto;padding:0 16px}
    .card{background:#121a2b;border:1px solid #223051;border-radius:12px;padding:18px;margin:14px 0}
    h1{font-size:26px;margin:0 0 14px}
    label{display:block;margin-bottom:8px;font-weight:600}
    input[type=file]{width:100%;margin-bottom:12px}
    .btn{display:inline-block;background:#2b6cff;border:none;color:#fff;padding:10px 14px;border-radius:10px;
         text-decoration:none;cursor:pointer;font-weight:700}
    .btn:hover{opacity:.9}
    .muted{opacity:.8}
    .row{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px}
    .err{background:#2a1a1a;border:1px solid #5b2a2a;padding:12px;border-radius:10px}
    code{background:#0a1020;padding:2px 6px;border-radius:6px}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>PDF → Excel (Bangla Voter List)</h1>

    <div class="card">
      <form action="/upload" method="post" enctype="multipart/form-data">
        <label>PDF Upload করুন:</label>
        <input type="file" name="pdf" accept=".pdf" required />
        <button class="btn" type="submit">Upload & Convert</button>
        <p class="muted">Upload হওয়ার পরে Excel অটো তৈরি হবে এবং নিচে Download বাটন আসবে।</p>
      </form>
    </div>

    {% if error %}
      <div class="card err">⚠️ {{ error }}</div>
    {% endif %}

    {% if success %}
      <div class="card">
        <h2 style="margin:0 0 10px">✅ Done!</h2>
        <p>Job ID: <b>{{ job_id }}</b></p>
        <p>Total rows: <b>{{ total }}</b></p>
        <div class="row">
          <a class="btn" href="/download/pdf/{{ pdf_file }}">⬇ Download PDF</a>
          <a class="btn" href="/download/excel/{{ excel_file }}">⬇ Download Excel</a>
        </div>
      </div>
    {% endif %}

    <div class="card">
      <b>Run (Local)</b>
      <div class="muted">
        1) Install: <code>pip install flask pandas openpyxl pymupdf pytesseract pillow</code><br/>
        2) Run: <code>python app.py</code><br/>
        3) Open: <code>http://127.0.0.1:5000</code><br/><br/>
        <b>Bangla OCR দরকার:</b> আপনার PDF-এ লেখা (cid:...) হলে OCR লাগবে।<br/>
        Ubuntu: <code>sudo apt install tesseract-ocr tesseract-ocr-ben</code>
      </div>
    </div>
  </div>
</body>
</html>
"""

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

def extract_text_from_pdf(pdf_path: str, force_ocr: bool = True) -> str:
    doc = fitz.open(pdf_path)
    chunks = []
    for page in doc:
        if force_ocr:
            t = ocr_page_to_text(page)
        else:
            t = page.get_text("text")
            if looks_garbled(t):
                t = ocr_page_to_text(page)
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

def pdf_to_excel(pdf_path: str, excel_out: str) -> int:
    # OCR ON by default (Bangla voter PDFs often garble)
    text = extract_text_from_pdf(pdf_path, force_ocr=True)
    records = parse_records(text)
    if not records:
        raise RuntimeError("No records found. OCR/format mismatch.")
    df = pd.DataFrame(records)
    df.to_excel(excel_out, index=False)
    return len(df)

# ======================
# ROUTES
# ======================
@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML, success=False)

@app.route("/upload", methods=["POST"])
def upload():
    if "pdf" not in request.files:
        return render_template_string(HTML, error="PDF ফাইল পাওয়া যায়নি!", success=False)

    f = request.files["pdf"]
    if not f.filename.lower().endswith(".pdf"):
        return render_template_string(HTML, error="শুধু PDF ফাইল আপলোড করুন!", success=False)

    job_id = str(uuid.uuid4())[:8]
    pdf_name = f"{job_id}.pdf"
    pdf_path = os.path.join(UPLOAD_DIR, pdf_name)
    f.save(pdf_path)

    excel_name = f"{job_id}.xlsx"
    excel_path = os.path.join(OUTPUT_DIR, excel_name)

    try:
        total = pdf_to_excel(pdf_path, excel_path)
    except Exception as e:
        return render_template_string(HTML, error=f"Convert ব্যর্থ হয়েছে: {e}", success=False)

    return render_template_string(
        HTML,
        success=True,
        job_id=job_id,
        total=total,
        pdf_file=pdf_name,
        excel_file=excel_name
    )

@app.route("/download/pdf/<filename>")
def download_pdf(filename):
    path = os.path.join(UPLOAD_DIR, filename)
    return send_file(path, as_attachment=True, download_name=filename)

@app.route("/download/excel/<filename>")
def download_excel(filename):
    path = os.path.join(OUTPUT_DIR, filename)
    return send_file(path, as_attachment=True, download_name=filename)

if __name__ == "__main__":
    # open http://127.0.0.1:5000
    app.run(host="0.0.0.0", port=5000, debug=True)
