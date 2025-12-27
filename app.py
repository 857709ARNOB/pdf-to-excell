import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
from convert_engine import pdf_to_excel

APP_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
OUTPUT_DIR = os.path.join(APP_DIR, "outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = "change-this-secret"

ALLOWED_EXT = {"pdf"}

def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "pdf" not in request.files:
        flash("PDF ফাইল সিলেক্ট করেননি!")
        return redirect(url_for("home"))

    f = request.files["pdf"]
    if f.filename == "":
        flash("PDF ফাইল সিলেক্ট করেননি!")
        return redirect(url_for("home"))

    if not allowed(f.filename):
        flash("শুধু PDF ফাইল আপলোড করুন!")
        return redirect(url_for("home"))

    job_id = str(uuid.uuid4())[:8]
    pdf_name = f"{job_id}.pdf"
    pdf_path = os.path.join(UPLOAD_DIR, pdf_name)
    f.save(pdf_path)

    excel_name = f"{job_id}.xlsx"
    excel_path = os.path.join(OUTPUT_DIR, excel_name)

    try:
        total = pdf_to_excel(pdf_path, excel_path, force_ocr=True)
    except Exception as e:
        flash(f"Convert ব্যর্থ হয়েছে: {e}")
        return redirect(url_for("home"))

    # success page with download links
    return render_template(
        "index.html",
        success=True,
        job_id=job_id,
        total=total,
        pdf_file=pdf_name,
        excel_file=excel_name
    )

@app.route("/download/pdf/<filename>")
def download_pdf(filename):
    return send_from_directory(UPLOAD_DIR, filename, as_attachment=True)

@app.route("/download/excel/<filename>")
def download_excel(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

if __name__ == "__main__":
    # Local run: http://127.0.0.1:5000
    app.run(host="0.0.0.0", port=5000, debug=True)
