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

# Rent calculator storage - uses Render persistent disk if available
# Set RENT_CALCULATOR_DIR env var to your Render disk mount (e.g., /var/data/rent_calculators)
RENT_CALCULATORS_DIR = os.environ.get('RENT_CALCULATOR_DIR', os.path.join(os.getcwd(), 'rent_calculators'))
ACTIVE_CALCULATOR_FILE = os.path.join(RENT_CALCULATORS_DIR, '.active')
os.makedirs(RENT_CALCULATORS_DIR, exist_ok=True)

# API Key for Excel Add-in authentication
# Set this in environment variable: AMI_OPTIX_API_KEY
API_KEY = os.environ.get('AMI_OPTIX_API_KEY', '')
# Admin key for rent calculator management (optional, defaults to API key)
ADMIN_KEY = os.environ.get('AMI_OPTIX_ADMIN_KEY', API_KEY)


def _validate_api_key():
    """Validate API key from request header. Returns error response or None if valid."""
    if not API_KEY:
        # No API key configured - allow all requests (dev mode)
        return None

    provided_key = request.headers.get('X-API-Key', '')
    if provided_key != API_KEY:
        return jsonify({"error": "Invalid or missing API key"}), 401
    return None


def _validate_admin_key():
    """Validate admin key for rent calculator management. Returns error response or None if valid."""
    if not ADMIN_KEY:
        # No admin key configured - allow all requests (dev mode)
        return None

    provided_key = request.headers.get('X-API-Key', '') or request.headers.get('X-Admin-Key', '')
    if provided_key != ADMIN_KEY:
        return jsonify({"error": "Invalid or missing admin key"}), 401
    return None


def _get_active_rent_calculator_path():
    """Get the path to the active rent calculator file."""
    # Check if there's an active selection
    if os.path.exists(ACTIVE_CALCULATOR_FILE):
        with open(ACTIVE_CALCULATOR_FILE, 'r') as f:
            active_name = f.read().strip()
        if active_name:
            active_path = os.path.join(RENT_CALCULATORS_DIR, active_name)
            if os.path.exists(active_path):
                return active_path

    # Fall back to default in repo root
    default_path = os.path.join(os.getcwd(), "2025 AMI Rent Calculator Unlocked.xlsx")
    if os.path.exists(default_path):
        return default_path

    return None


def _list_rent_calculators():
    """List all available rent calculator files."""
    calculators = []
    active_name = None

    # Get active calculator name
    if os.path.exists(ACTIVE_CALCULATOR_FILE):
        with open(ACTIVE_CALCULATOR_FILE, 'r') as f:
            active_name = f.read().strip()

    # List uploaded calculators
    if os.path.exists(RENT_CALCULATORS_DIR):
        for filename in os.listdir(RENT_CALCULATORS_DIR):
            if filename.startswith('.'):
                continue
            if filename.lower().endswith(('.xlsx', '.xlsm')):
                filepath = os.path.join(RENT_CALCULATORS_DIR, filename)
                stat = os.stat(filepath)
                calculators.append({
                    'name': filename,
                    'size': stat.st_size,
                    'modified': stat.st_mtime,
                    'is_active': filename == active_name,
                    'source': 'uploaded'
                })

    # Add default calculator if it exists
    default_path = os.path.join(os.getcwd(), "2025 AMI Rent Calculator Unlocked.xlsx")
    if os.path.exists(default_path):
        stat = os.stat(default_path)
        is_default_active = not active_name  # Default is active if no selection
        calculators.append({
            'name': '2025 AMI Rent Calculator Unlocked.xlsx',
            'size': stat.st_size,
            'modified': stat.st_mtime,
            'is_active': is_default_active,
            'source': 'default'
        })

    return calculators


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
        rent_calc_path = _get_active_rent_calculator_path()
        if rent_calc_path:
            try:
                rent_schedule = load_rent_schedule(rent_calc_path)
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


# =============================================================================
# RENT CALCULATOR MANAGEMENT ENDPOINTS
# =============================================================================

@app.route('/api/rent-calculators', methods=['GET'])
def list_rent_calculators():
    """List all available rent calculator files."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    calculators = _list_rent_calculators()
    active_path = _get_active_rent_calculator_path()

    return jsonify({
        "calculators": calculators,
        "active_path": active_path,
        "storage_dir": RENT_CALCULATORS_DIR
    })


@app.route('/api/rent-calculators/upload', methods=['POST'])
def upload_rent_calculator():
    """Upload a new rent calculator file."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    # Validate file extension
    filename = secure_filename(file.filename)
    if not filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({"error": "Only .xlsx and .xlsm files are supported"}), 400

    # Save the file
    filepath = os.path.join(RENT_CALCULATORS_DIR, filename)

    # Check if file already exists
    if os.path.exists(filepath):
        overwrite = request.form.get('overwrite', 'false').lower() == 'true'
        if not overwrite:
            return jsonify({"error": f"File '{filename}' already exists. Set overwrite=true to replace."}), 409

    try:
        file.save(filepath)

        # Validate the file is a valid rent calculator by trying to load it
        try:
            schedule = load_rent_schedule(filepath)
            # Basic validation - check it has rent data
            if not schedule.gross_rents:
                os.remove(filepath)
                return jsonify({"error": "File does not appear to be a valid rent calculator (no rent data found)"}), 400
        except Exception as e:
            os.remove(filepath)
            return jsonify({"error": f"Invalid rent calculator file: {str(e)}"}), 400

        return jsonify({
            "success": True,
            "message": f"Uploaded {filename}",
            "filename": filename
        })

    except Exception as e:
        return jsonify({"error": f"Upload failed: {str(e)}"}), 500


@app.route('/api/rent-calculators/activate', methods=['POST'])
def activate_rent_calculator():
    """Set the active rent calculator."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    data = request.get_json()
    if not data or 'name' not in data:
        return jsonify({"error": "Missing 'name' in request body"}), 400

    name = data['name']

    # Check if it's the default (clear the active selection)
    if name == '2025 AMI Rent Calculator Unlocked.xlsx' or name == 'default':
        if os.path.exists(ACTIVE_CALCULATOR_FILE):
            os.remove(ACTIVE_CALCULATOR_FILE)
        return jsonify({
            "success": True,
            "message": "Activated default rent calculator",
            "active": "2025 AMI Rent Calculator Unlocked.xlsx"
        })

    # Verify the file exists
    filepath = os.path.join(RENT_CALCULATORS_DIR, name)
    if not os.path.exists(filepath):
        return jsonify({"error": f"Rent calculator '{name}' not found"}), 404

    # Set as active
    with open(ACTIVE_CALCULATOR_FILE, 'w') as f:
        f.write(name)

    return jsonify({
        "success": True,
        "message": f"Activated {name}",
        "active": name
    })


@app.route('/api/rent-calculators/<filename>', methods=['DELETE'])
def delete_rent_calculator(filename):
    """Delete a rent calculator file."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    # Cannot delete the default
    if filename == '2025 AMI Rent Calculator Unlocked.xlsx':
        return jsonify({"error": "Cannot delete the default rent calculator"}), 400

    filepath = os.path.join(RENT_CALCULATORS_DIR, secure_filename(filename))
    if not os.path.exists(filepath):
        return jsonify({"error": f"File '{filename}' not found"}), 404

    # If this was the active calculator, clear the selection
    if os.path.exists(ACTIVE_CALCULATOR_FILE):
        with open(ACTIVE_CALCULATOR_FILE, 'r') as f:
            active_name = f.read().strip()
        if active_name == filename:
            os.remove(ACTIVE_CALCULATOR_FILE)

    os.remove(filepath)

    return jsonify({
        "success": True,
        "message": f"Deleted {filename}"
    })


@app.route('/admin/rent-calculators')
def rent_calculator_admin():
    """Simple admin page for rent calculator management."""
    html = """
<!DOCTYPE html>
<html>
<head>
    <title>AMI Optix - Rent Calculator Admin</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 800px; margin: 40px auto; padding: 20px; }
        h1 { color: #2c3e50; }
        .card { background: #f8f9fa; border-radius: 8px; padding: 20px; margin: 20px 0; }
        .active { background: #d4edda; border-left: 4px solid #28a745; }
        .default { background: #e7f3ff; border-left: 4px solid #007bff; }
        button { background: #007bff; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; margin-right: 8px; }
        button:hover { background: #0056b3; }
        button.danger { background: #dc3545; }
        button.danger:hover { background: #c82333; }
        button.success { background: #28a745; }
        input[type="file"] { margin: 10px 0; }
        .status { padding: 10px; border-radius: 4px; margin: 10px 0; }
        .status.error { background: #f8d7da; color: #721c24; }
        .status.success { background: #d4edda; color: #155724; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; }
        .badge.active { background: #28a745; color: white; }
        .badge.default { background: #007bff; color: white; }
    </style>
</head>
<body>
    <h1>🏠 Rent Calculator Admin</h1>

    <div class="card">
        <h2>Upload New Calculator</h2>
        <p>Upload a new rent calculator Excel file (.xlsx or .xlsm)</p>
        <input type="file" id="fileInput" accept=".xlsx,.xlsm">
        <br>
        <label><input type="checkbox" id="overwrite"> Overwrite if exists</label>
        <br><br>
        <button onclick="uploadFile()">Upload</button>
        <div id="uploadStatus"></div>
    </div>

    <div class="card">
        <h2>Available Calculators</h2>
        <div id="calculatorList">Loading...</div>
    </div>

    <script>
        const apiKey = prompt('Enter API Key:');

        async function fetchCalculators() {
            try {
                const res = await fetch('/api/rent-calculators', {
                    headers: { 'X-API-Key': apiKey }
                });
                const data = await res.json();
                if (data.error) {
                    document.getElementById('calculatorList').innerHTML = '<p class="status error">' + data.error + '</p>';
                    return;
                }
                renderCalculators(data.calculators);
            } catch (e) {
                document.getElementById('calculatorList').innerHTML = '<p class="status error">Failed to load: ' + e.message + '</p>';
            }
        }

        function renderCalculators(calculators) {
            if (!calculators || calculators.length === 0) {
                document.getElementById('calculatorList').innerHTML = '<p>No calculators found.</p>';
                return;
            }

            let html = '<table><tr><th>Name</th><th>Source</th><th>Size</th><th>Actions</th></tr>';
            calculators.forEach(calc => {
                const badges = [];
                if (calc.is_active) badges.push('<span class="badge active">Active</span>');
                if (calc.source === 'default') badges.push('<span class="badge default">Default</span>');

                const actions = [];
                if (!calc.is_active) {
                    actions.push('<button onclick="activateCalc(\\'' + calc.name + '\\')">Activate</button>');
                }
                if (calc.source !== 'default') {
                    actions.push('<button class="danger" onclick="deleteCalc(\\'' + calc.name + '\\')">Delete</button>');
                }

                html += '<tr class="' + (calc.is_active ? 'active' : '') + ' ' + (calc.source === 'default' ? 'default' : '') + '">';
                html += '<td>' + calc.name + ' ' + badges.join(' ') + '</td>';
                html += '<td>' + calc.source + '</td>';
                html += '<td>' + Math.round(calc.size / 1024) + ' KB</td>';
                html += '<td>' + actions.join('') + '</td>';
                html += '</tr>';
            });
            html += '</table>';
            document.getElementById('calculatorList').innerHTML = html;
        }

        async function uploadFile() {
            const fileInput = document.getElementById('fileInput');
            const overwrite = document.getElementById('overwrite').checked;
            const statusDiv = document.getElementById('uploadStatus');

            if (!fileInput.files[0]) {
                statusDiv.innerHTML = '<p class="status error">Please select a file</p>';
                return;
            }

            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            formData.append('overwrite', overwrite);

            statusDiv.innerHTML = '<p>Uploading...</p>';

            try {
                const res = await fetch('/api/rent-calculators/upload', {
                    method: 'POST',
                    headers: { 'X-API-Key': apiKey },
                    body: formData
                });
                const data = await res.json();
                if (data.error) {
                    statusDiv.innerHTML = '<p class="status error">' + data.error + '</p>';
                } else {
                    statusDiv.innerHTML = '<p class="status success">' + data.message + '</p>';
                    fetchCalculators();
                }
            } catch (e) {
                statusDiv.innerHTML = '<p class="status error">Upload failed: ' + e.message + '</p>';
            }
        }

        async function activateCalc(name) {
            try {
                const res = await fetch('/api/rent-calculators/activate', {
                    method: 'POST',
                    headers: {
                        'X-API-Key': apiKey,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ name: name })
                });
                const data = await res.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                } else {
                    fetchCalculators();
                }
            } catch (e) {
                alert('Failed: ' + e.message);
            }
        }

        async function deleteCalc(name) {
            if (!confirm('Delete ' + name + '?')) return;

            try {
                const res = await fetch('/api/rent-calculators/' + encodeURIComponent(name), {
                    method: 'DELETE',
                    headers: { 'X-API-Key': apiKey }
                });
                const data = await res.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                } else {
                    fetchCalculators();
                }
            } catch (e) {
                alert('Failed: ' + e.message);
            }
        }

        fetchCalculators();
    </script>
</body>
</html>
"""
    return html, 200, {'Content-Type': 'text/html'}


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

