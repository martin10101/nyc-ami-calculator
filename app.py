import os
import shutil
import tempfile
import zipfile
import json
import math
import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from main import main as run_ami_optix_analysis, default_converter
from ami_optix.narrator import generate_internal_summary
from ami_optix.report_generator import create_excel_reports
from ami_optix.config_loader import load_config
from ami_optix.solver import find_optimal_scenarios
from ami_optix.rent_calculator import load_rent_schedule, compute_rents_for_assignments

app = Flask(__name__)
UPLOADS_DIR = os.path.join(os.getcwd(), 'uploads')
DASHBOARD_DIR = os.path.join(os.getcwd(), 'dashboard_static')
os.makedirs(UPLOADS_DIR, exist_ok=True)

# API Key for Excel Add-in authentication
# Set this in environment variable: AMI_OPTIX_API_KEY
API_KEY = os.environ.get('AMI_OPTIX_API_KEY', '')


def _validate_api_key():
    """Validate API key from request header. Returns error response or None if valid."""
    if not API_KEY:
        # No API key configured - allow all requests (dev mode)
        return None

    provided_key = request.headers.get('X-API-Key', '')
    if provided_key != API_KEY:
        return jsonify({"error": "Invalid or missing API key"}), 401
    return None


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


@app.route('/api/optimize', methods=['POST'])
def optimize_units():
    """
    JSON-based optimization endpoint for Excel VBA Add-in.

    Accepts unit data directly as JSON (no file upload needed).
    Requires API key authentication via X-API-Key header.

    Request body:
    {
        "units": [
            {"unit_id": "1A", "bedrooms": 2, "net_sf": 850, "floor": 1},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 650, "floor": 2}
        ],
        "utilities": {
            "electricity": "tenant_pays",
            "cooking": "gas",
            "heat": "gas",
            "hot_water": "gas"
        }
    }

    Returns:
    {
        "scenarios": { ... },
        "notes": [ ... ],
        "project_summary": { ... }
    }
    """
    # Validate API key
    auth_error = _validate_api_key()
    if auth_error:
        return auth_error

    # Parse JSON body
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "Request body must be JSON"}), 400
    except Exception as e:
        return jsonify({"error": f"Invalid JSON: {str(e)}"}), 400

    # Extract units
    units = data.get('units', [])
    if not units or not isinstance(units, list):
        return jsonify({"error": "Missing or invalid 'units' array"}), 400

    # Validate required fields for each unit
    required_fields = ['unit_id', 'bedrooms', 'net_sf']
    for i, unit in enumerate(units):
        for field in required_fields:
            if field not in unit:
                return jsonify({"error": f"Unit {i+1} missing required field: {field}"}), 400

    # Extract utilities (with defaults)
    utilities = data.get('utilities', {})
    utilities_clean = {
        'electricity': utilities.get('electricity', 'na'),
        'cooking': utilities.get('cooking', 'na'),
        'heat': utilities.get('heat', 'na'),
        'hot_water': utilities.get('hot_water', 'na'),
    }

    try:
        # Convert units to DataFrame (same format parser produces)
        df_units = pd.DataFrame(units)

        # Ensure required columns exist with correct types
        df_units['unit_id'] = df_units['unit_id'].astype(str)
        df_units['bedrooms'] = pd.to_numeric(df_units['bedrooms'], errors='coerce')
        df_units['net_sf'] = pd.to_numeric(df_units['net_sf'], errors='coerce')

        # Optional columns
        if 'floor' in df_units.columns:
            df_units['floor'] = pd.to_numeric(df_units['floor'], errors='coerce')
        if 'balcony' in df_units.columns:
            df_units['balcony'] = df_units['balcony'].astype(bool)

        # Add client_ami column (required by solver, will be overwritten)
        # Set to a placeholder that indicates "needs assignment"
        if 'client_ami' not in df_units.columns:
            df_units['client_ami'] = 0.6  # Default placeholder

        # Load config
        config = load_config()

        # Run solver
        solver_results = find_optimal_scenarios(df_units, config)
        scenarios = solver_results.get('scenarios', {})
        notes = solver_results.get('notes', [])

        if not scenarios or not scenarios.get('absolute_best'):
            return jsonify({
                "success": False,
                "error": "No optimal solution found",
                "notes": notes
            }), 200

        # Load rent schedule and apply rent calculations
        rent_schedule = None
        default_rent_path = os.path.join(os.getcwd(), "2025 AMI Rent Calculator Unlocked.xlsx")
        if os.path.exists(default_rent_path):
            try:
                rent_schedule = load_rent_schedule(default_rent_path)
            except Exception as e:
                notes.append(f"Warning: Could not load rent calculator: {str(e)}")

        # Apply rent calculations to each scenario
        if rent_schedule:
            for scenario_key, scenario in scenarios.items():
                if scenario and 'assignments' in scenario:
                    assignments, rent_totals = compute_rents_for_assignments(
                        rent_schedule,
                        scenario['assignments'],
                        utilities_clean
                    )
                    scenario['assignments'] = assignments
                    scenario['rent_totals'] = rent_totals

        # Build response
        response = {
            "success": True,
            "scenarios": _sanitize_for_json(scenarios),
            "notes": notes,
            "project_summary": {
                "total_units": len(df_units),
                "total_sf": float(df_units['net_sf'].sum()),
                "utility_selections": utilities_clean
            }
        }

        return jsonify(response)

    except Exception as e:
        app.logger.exception("optimize_units failed: %s", e)
        return jsonify({"error": f"Optimization failed: {str(e)}"}), 500


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

            utilities_payload = None
            overrides_payload = None
            rent_calculator_path = None

            utilities_raw = request.form.get('utilities')
            if utilities_raw:
                try:
                    utilities_payload = json.loads(utilities_raw)
                except json.JSONDecodeError:
                    return jsonify({"error": "Invalid utilities payload."}), 400

            overrides_raw = request.form.get('overrides')
            if overrides_raw:
                try:
                    overrides_payload = json.loads(overrides_raw)
                except json.JSONDecodeError:
                    return jsonify({"error": "Invalid overrides payload."}), 400

            rent_calculator_upload = request.files.get('rentCalculator')
            if rent_calculator_upload and rent_calculator_upload.filename:
                rent_filename = secure_filename(rent_calculator_upload.filename)
                rent_calculator_path = os.path.join(temp_dir, rent_filename)
                rent_calculator_upload.save(rent_calculator_path)

            try:
                # 1. Run the core analysis
                analysis_output = run_ami_optix_analysis(
                    upload_filepath,
                    utilities=utilities_payload,
                    overrides=overrides_payload,
                    rent_calculator_path=rent_calculator_path,
                )
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
                prefer_xlsb = filename.lower().endswith('.xlsb') or bool(request.form.get('preferXlsb'))
                utility_selections = analysis_results.get('project_summary', {}).get('utility_selections')
                rent_workbook_info = analysis_results.get('rent_workbook') or {}
                rent_workbook_source = rent_workbook_info.get('source_path')

                report_files = create_excel_reports(
                    analysis_results,
                    upload_filepath,
                    original_headers,
                    output_dir=temp_dir,
                    prefer_xlsb=prefer_xlsb,
                    utilities=utility_selections,
                    rent_workbook_path=rent_workbook_source,
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

