import fitz  # PyMuPDF
import pandas as pd
import re

PDF_FILE = "261066_com_6160_female_without_photo_347_2025-11-24.pdf"
OUTPUT_FILE = "output.xlsx"

bn2en = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")

def normalize(s):
    return s.translate(bn2en)

doc = fitz.open(PDF_FILE)

rows = []
current = None

start_pat = re.compile(r"^(\d{4})\.\s*নাম:\s*(.+)$")

def save_current():
    global current
    if current and current["ভোটার নং"]:
        rows.append(current)
    current = None

for page in doc:
    lines = [l.strip() for l in page.get_text().splitlines() if l.strip()]
    for line in lines:
        ln = normalize(line)

        m = start_pat.match(ln)
        if m:
            save_current()
            current = {
                "Serial": int(m.group(1)),
                "নাম": m.group(2),
                "ভোটার নং": "",
                "পিতা": "",
                "মাতা": "",
                "পেশা": "",
                "জন্ম তারিখ": "",
                "ঠিকানা": ""
            }
            continue

        if not current:
            continue

        if "মাইগ্রেট" in ln:
            current = None
            continue

        if ln.startswith("ভোটার নং:"):
            current["ভোটার নং"] = ln.split(":",1)[1].strip()
        elif ln.startswith("পিতা:"):
            current["পিতা"] = ln.split(":",1)[1].strip()
        elif ln.startswith("মাতা:"):
            current["মাতা"] = ln.split(":",1)[1].strip()
        elif ln.startswith("পেশা:"):
            rest = ln.split(":",1)[1]
            if ",জন্ম তারিখ:" in rest:
                p, d = rest.split(",জন্ম তারিখ:")
                current["পেশা"] = p.strip()
                current["জন্ম তারিখ"] = d.strip()
            else:
                current["পেশা"] = rest.strip()
        elif ln.startswith("ঠিকানা:"):
            current["ঠিকানা"] = ln.split(":",1)[1].strip()

save_current()

df = pd.DataFrame(rows).sort_values("Serial")
df.to_excel(OUTPUT_FILE, index=False)

print("✅ Excel তৈরি হয়েছে:", OUTPUT_FILE)
