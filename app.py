import os
import shutil
import tempfile
import zipfile
import json
import math
import numpy as np
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from main import main as run_ami_optix_analysis, default_converter
from ami_optix.narrator import generate_internal_summary
from ami_optix.report_generator import create_excel_reports

app = Flask(__name__)
UPLOADS_DIR = os.path.join(os.getcwd(), 'uploads')
DASHBOARD_DIR = os.path.join(os.getcwd(), 'dashboard_static')
os.makedirs(UPLOADS_DIR, exist_ok=True)


def _sanitize_for_json(value):
    """Recursively convert numpy/pandas types and NaN values to JSON-safe primitives."""
    if isinstance(value, dict):
        return {key: _sanitize_for_json(val) for key, val in value.items()}
    if isinstance(value, list):
        return [_sanitize_for_json(item) for item in value]
    if isinstance(value, tuple):
        return [_sanitize_for_json(item) for item in value]
    if isinstance(value, np.ndarray):
        return [_sanitize_for_json(item) for item in value.tolist()]
    if isinstance(value, (float, np.floating)):
        if math.isnan(value) or math.isinf(value):
            return None
        return float(value)
    if isinstance(value, (np.integer,)):
        return int(value)
    try:
        return default_converter(value)
    except TypeError:
        return value


def _dashboard_file_exists(filename: str) -> bool:
    return os.path.exists(os.path.join(DASHBOARD_DIR, filename))


@app.route('/healthz')
def healthcheck():
    """Lightweight health endpoint for uptime checks."""
    return jsonify({"status": "ok"})


@app.route('/api/analyze', methods=['POST'])
def analyze_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected for uploading"}), 400

    if file:
        filename = secure_filename(file.filename)

        # Use a temporary directory for all processing
        with tempfile.TemporaryDirectory() as temp_dir:
            upload_filepath = os.path.join(temp_dir, filename)
            file.save(upload_filepath)

            try:
                # 1. Run the core analysis
                analysis_output = run_ami_optix_analysis(upload_filepath)
                if "error" in analysis_output:
                    return jsonify(analysis_output), 400

                analysis_results = analysis_output['results']
                original_headers = analysis_output['original_headers']
                analysis_meta = analysis_results.get('analysis_meta', {})
                app.logger.info("analysis_id=%s combos=%s unique=%s duration=%.2fs truncated=%s", analysis_meta.get('analysis_id'), analysis_meta.get('solver_combination_count'), analysis_meta.get('solver_unique_scenarios'), analysis_meta.get('duration_sec'), analysis_meta.get('truncated'))

                # 2. Generate the internal summary (no LLM for now)
                narrative = generate_internal_summary(analysis_results)
                analysis_results['narrative_analysis'] = narrative

                # 3. Generate Excel reports, passing the original headers
                report_files = create_excel_reports(
                    analysis_results,
                    upload_filepath,
                    original_headers,
                    output_dir=temp_dir
                )

                # 4. Create a zip file containing all reports
                zip_filename = f"{os.path.splitext(filename)[0]}_reports.zip"
                zip_filepath = os.path.join(temp_dir, zip_filename)
                with zipfile.ZipFile(zip_filepath, 'w') as zipf:
                    for report_file in report_files:
                        zipf.write(report_file, os.path.basename(report_file))

                # 5. Add a download link for the zip file to the response
                analysis_results['download_link'] = f"/api/download/{zip_filename}"

                # Persist the zip in the uploads directory for the download endpoint
                os.makedirs(UPLOADS_DIR, exist_ok=True)
                shutil.move(zip_filepath, os.path.join(UPLOADS_DIR, zip_filename))

                safe_payload = _sanitize_for_json(analysis_results)

                return jsonify(safe_payload)

            except Exception as e:
                app.logger.exception("analysis_failed: %s", e)
                return jsonify({"error": f"An unexpected error occurred during analysis: {str(e)}"}), 500

    return jsonify({"error": "An unknown error occurred"}), 500


@app.route('/api/download/<filename>', methods=['GET'])
def download_report(filename):
    """Serves the generated zip file for download."""
    try:
        return send_from_directory(UPLOADS_DIR, filename, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found."}), 404


@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_dashboard(path):
    """Serve the static dashboard that ships with the deployment."""
    # Guard API routes so they keep flowing to their own handlers
    if path.startswith('api/'):
        return jsonify({"error": "Not found"}), 404

    requested = path or 'index.html'
    requested_path = os.path.join(DASHBOARD_DIR, requested)

    if os.path.isdir(requested_path):
        requested = os.path.join(requested, 'index.html')

    if _dashboard_file_exists(requested):
        return send_from_directory(DASHBOARD_DIR, requested)

    if _dashboard_file_exists('index.html'):
        # SPA fallback - return index so client-side routing can take over
        return send_from_directory(DASHBOARD_DIR, 'index.html')

    return (
        "<html><head><title>NYC AMI Calculator API</title></head><body>"
        "<h1>NYC AMI Calculator API</h1>"
        "<p>The interactive dashboard has not been built. "
        "Deploy tip: run 'npm install' and 'npm run build && npm run export' inside the"
        " dashboard/ folder, then deploy the contents of the generated dashboard_static/"
        " directory alongside this service.</p>"
        "</body></html>",
        200,
        {"Content-Type": "text/html; charset=utf-8"}
    )


if __name__ == '__main__':
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    app.run(debug=True, port=5001)

