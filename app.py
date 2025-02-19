from flask import Flask, render_template, request, send_file, flash
import os
from ubah import main  # Skrip pemrosesan PDF ke Excel

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Dibutuhkan untuk flash messages

# Folder tujuan untuk menyimpan file yang diupload
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")  # Gantilah sesuai kebutuhan
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Hanya izinkan file dengan ekstensi .pdf
ALLOWED_EXTENSIONS = {"pdf"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            flash("❌ Tidak ada file yang diupload!", "danger")
            return render_template("index.html")

        pdf_file = request.files["file"]

        if pdf_file.filename == "":
            flash("⚠️ Pilih file terlebih dahulu!", "warning")
            return render_template("index.html")

        if not allowed_file(pdf_file.filename):
            flash("🚫 Format file tidak diizinkan! Hanya PDF yang diperbolehkan.", "danger")
            return render_template("index.html")

        # Simpan file ke folder tujuan
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], pdf_file.filename)
        pdf_file.save(file_path)

        # Proses PDF
        main(file_path)
        output_file = file_path.replace(".pdf", ".xlsx")

        return send_file(output_file, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=8080, debug=True)
