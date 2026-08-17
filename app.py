from __future__ import annotations

import shutil
import tempfile
import threading
import time
import uuid
from pathlib import Path

from flask import Flask, abort, render_template_string, request, send_file

from catalog_exporter import TEMPLATE_FORM_VARIANTS, export_catalog

APP_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = APP_DIR / "uploads"
UPLOADS_DIR.mkdir(exist_ok=True)

MAX_CONTENT_LENGTH = 80 * 1024 * 1024  # 80 MB per upload
JOB_MAX_AGE_SECONDS = 3600  # sweep anything older than 1 hour
DELAYED_CLEANUP_SECONDS = 10  # grace period so send_file can finish streaming

TEMPLATE_FORM_CHOICES = ["auto", *TEMPLATE_FORM_VARIANTS.keys()]

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

PAGE = r"""
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Catalog PDF Exporter</title>
<style>
  :root { color-scheme: light dark; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    max-width: 480px; margin: 0 auto; padding: 24px 16px 60px;
    background: #0f1115; color: #e8e8e8;
  }
  h1 { font-size: 1.3rem; margin-bottom: 4px; }
  p.sub { color: #9aa0a6; margin-top: 0; font-size: 0.9rem; }
  form { display: flex; flex-direction: column; gap: 18px; margin-top: 24px; }
  label { font-weight: 600; font-size: 0.95rem; display: block; margin-bottom: 6px; }
  input[type=file] {
    width: 100%; padding: 10px; border-radius: 10px; border: 1px solid #333;
    background: #1a1d24; color: #e8e8e8; font-size: 0.9rem;
  }
  select {
    width: 100%; padding: 12px; border-radius: 10px; border: 1px solid #333;
    background: #1a1d24; color: #e8e8e8; font-size: 1rem;
  }
  button {
    padding: 16px; border-radius: 12px; border: none; font-size: 1.05rem;
    font-weight: 600; background: #4f7cff; color: white; cursor: pointer;
  }
  button:disabled { background: #33415e; }
  .field { background: #171a21; padding: 14px; border-radius: 14px; }
  .error {
    background: #3a1c1c; border: 1px solid #7a2e2e; color: #ff9a9a;
    padding: 12px 14px; border-radius: 10px; margin-top: 18px; font-size: 0.9rem;
  }
  #status { margin-top: 16px; font-size: 0.9rem; color: #9aa0a6; display: none; }
  #status.done { color: #7ad48c; }
  #status.failed { color: #ff9a9a; }
</style>
</head>
<body>
  <h1>Catalog PDF Exporter</h1>
  <p class="sub">Upload this week's Excel + template PDF, get the finished catalog back.</p>

  {% if error %}
  <div class="error">{{ error }}</div>
  {% endif %}

  <form id="genForm" method="post" action="/generate" enctype="multipart/form-data">
    <div class="field">
      <label for="excel">Excel file (.xlsx)</label>
      <input type="file" id="excel" name="excel" accept=".xlsx" required>
    </div>
    <div class="field">
      <label for="template">Template PDF (.pdf)</label>
      <input type="file" id="template" name="template" accept=".pdf" required>
    </div>
    <div class="field">
      <label for="template_form">Template form</label>
      <select id="template_form" name="template_form">
        {% for choice in template_forms %}
        <option value="{{ choice }}" {% if choice == 'auto' %}selected{% endif %}>{{ choice }}</option>
        {% endfor %}
      </select>
    </div>
    <div class="field">
      <label for="quality">Quality</label>
      <select id="quality" name="quality">
        <option value="normal" selected>normal</option>
        <option value="high">high</option>
      </select>
    </div>
    <button type="submit" id="submitBtn">Generate PDF</button>
    <div id="status">Generating... this can take a minute, please keep this tab open.</div>
  </form>

  <script>
    document.getElementById('genForm').addEventListener('submit', async function (event) {
      event.preventDefault();

      const form = event.target;
      const btn = document.getElementById('submitBtn');
      const status = document.getElementById('status');

      btn.disabled = true;
      btn.textContent = 'Generating...';
      status.className = '';
      status.style.display = 'block';
      status.textContent = 'Generating... this can take a minute, please keep this tab open.';

      try {
        const response = await fetch(form.action, {
          method: 'POST',
          body: new FormData(form),
        });

        if (!response.ok) {
          const contentType = response.headers.get('Content-Type') || '';
          let message = `Export failed (HTTP ${response.status}).`;
          if (contentType.includes('text/html')) {
            const html = await response.text();
            const match = html.match(/<div class="error">([\s\S]*?)<\/div>/);
            if (match) message = match[1].trim();
          }
          status.textContent = message;
          status.className = 'failed';
          return;
        }

        const blob = await response.blob();
        const disposition = response.headers.get('Content-Disposition') || '';
        const nameMatch = disposition.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/);
        const filename = nameMatch ? decodeURIComponent(nameMatch[1]) : 'output.pdf';

        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);

        status.textContent = 'Done — PDF downloaded.';
        status.className = 'done';
      } catch (err) {
        status.textContent = 'Export failed: network error. Please try again.';
        status.className = 'failed';
      } finally {
        btn.disabled = false;
        btn.textContent = 'Generate PDF';
      }
    });
  </script>
</body>
</html>
"""


def cleanup_old_jobs() -> None:
    cutoff = time.time() - JOB_MAX_AGE_SECONDS
    for child in UPLOADS_DIR.iterdir():
        try:
            if child.is_dir() and child.stat().st_mtime < cutoff:
                shutil.rmtree(child, ignore_errors=True)
        except FileNotFoundError:
            continue


def schedule_cleanup(job_dir: Path) -> None:
    def _cleanup() -> None:
        time.sleep(DELAYED_CLEANUP_SECONDS)
        shutil.rmtree(job_dir, ignore_errors=True)

    threading.Thread(target=_cleanup, daemon=True).start()


@app.route("/", methods=["GET"])
def index():
    cleanup_old_jobs()
    return render_template_string(PAGE, template_forms=TEMPLATE_FORM_CHOICES, error=None)


@app.route("/generate", methods=["POST"])
def generate():
    cleanup_old_jobs()

    excel_file = request.files.get("excel")
    template_file = request.files.get("template")
    quality = request.form.get("quality", "normal")
    template_form = request.form.get("template_form", "auto")

    if not excel_file or not excel_file.filename.lower().endswith(".xlsx"):
        return render_template_string(
            PAGE, template_forms=TEMPLATE_FORM_CHOICES, error="Please upload a valid .xlsx Excel file."
        ), 400
    if not template_file or not template_file.filename.lower().endswith(".pdf"):
        return render_template_string(
            PAGE, template_forms=TEMPLATE_FORM_CHOICES, error="Please upload a valid .pdf template file."
        ), 400
    if quality not in ("normal", "high"):
        abort(400, "Invalid quality.")
    if template_form not in TEMPLATE_FORM_CHOICES:
        abort(400, "Invalid template form.")

    job_id = uuid.uuid4().hex
    job_dir = UPLOADS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    excel_path = job_dir / "input.xlsx"
    template_path = job_dir / "template.pdf"
    excel_file.save(excel_path)
    template_file.save(template_path)

    output_stem = Path(excel_file.filename).stem or "catalog"
    output_path = job_dir / "output.pdf"

    try:
        export_catalog(
            excel_path=excel_path,
            template_pdf=template_path,
            output_path=output_path,
            quality=quality,
            template_form=template_form,
        )
    except Exception as exc:  # noqa: BLE001 - surface any export failure to the user
        shutil.rmtree(job_dir, ignore_errors=True)
        return render_template_string(
            PAGE, template_forms=TEMPLATE_FORM_CHOICES, error=f"Export failed: {exc}"
        ), 500

    response = send_file(output_path, as_attachment=True, download_name=f"{output_stem} output.pdf")
    schedule_cleanup(job_dir)
    return response


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8020, debug=False)
