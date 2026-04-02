import os
import uuid
import shutil
from flask import Flask, render_template, request, send_file, after_this_request
from werkzeug.utils import secure_filename
from converter import convert_docx_to_qti, parse_mark_scheme

app = Flask(__name__)

MAX_SIZE = 10 * 1024 * 1024  # 10 MB
app.config['MAX_CONTENT_LENGTH'] = MAX_SIZE

BASE_UPLOAD = "jobs"
os.makedirs(BASE_UPLOAD, exist_ok=True)


def allowed_file(filename):
    return filename.lower().endswith(".docx")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    # ── Question paper (required) ──────────────────────────────────────
    if "file" not in request.files:
        return {"error": "No file part in request"}, 400

    file = request.files["file"]

    if file.filename == "":
        return {"error": "No file selected"}, 400

    if not allowed_file(file.filename):
        return {"error": "Only .docx files are allowed"}, 400

    # ── Create per-job working directory ──────────────────────────────
    job_id  = str(uuid.uuid4())
    job_dir = os.path.join(BASE_UPLOAD, job_id)
    os.makedirs(job_dir)

    # Save question paper
    safe_name  = secure_filename(file.filename)
    input_path = os.path.join(job_dir, safe_name)
    file.save(input_path)

    if os.path.getsize(input_path) > MAX_SIZE:
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"error": "File exceeds 10 MB limit"}, 400

    # ── Mark scheme (optional) ─────────────────────────────────────────
    ms_answers = {}
    ms_file    = request.files.get("ms_file")

    if ms_file and ms_file.filename and allowed_file(ms_file.filename):
        ms_safe   = secure_filename(ms_file.filename)
        ms_path   = os.path.join(job_dir, ms_safe)
        ms_file.save(ms_path)

        if os.path.getsize(ms_path) <= MAX_SIZE:
            try:
                ms_answers = parse_mark_scheme(ms_path)
            except Exception:
                ms_answers = {}   # ignore MS errors — still convert without answers

    # ── Convert ────────────────────────────────────────────────────────
    try:
        zip_path = convert_docx_to_qti(input_path, job_dir, ms_answers=ms_answers)
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"error": str(e)}, 500

    @after_this_request
    def cleanup(response):
        try:
            response.call_on_close(lambda: shutil.rmtree(job_dir, ignore_errors=True))
        except Exception:
            shutil.rmtree(job_dir, ignore_errors=True)
        return response

    # Use the original filename (sans .docx) as the download name
    base_name = os.path.splitext(file.filename)[0]
    download_name = base_name + ".zip"

    return send_file(
        zip_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/zip"
    )


@app.route("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False)
