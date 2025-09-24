import os
import tempfile
import zipfile
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from main import main as run_ami_optix_analysis
from ami_optix.narrator import generate_internal_summary
from ami_optix.report_generator import create_excel_reports

app = Flask(__name__)

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

                # 2. Generate the internal summary (no LLM for now)
                narrative = generate_internal_summary(analysis_results)
                analysis_results['narrative_analysis'] = narrative

                # 3. Generate Excel reports, passing the original headers
                report_files = create_excel_reports(analysis_results, upload_filepath, original_headers, output_dir=temp_dir)

                # 4. Create a zip file containing all reports
                zip_filename = f"{os.path.splitext(filename)[0]}_reports.zip"
                zip_filepath = os.path.join(temp_dir, zip_filename)
                with zipfile.ZipFile(zip_filepath, 'w') as zipf:
                    for report_file in report_files:
                        zipf.write(report_file, os.path.basename(report_file))

                # 5. Add a download link for the zip file to the response
                analysis_results['download_link'] = f"/api/download/{zip_filename}"

                # Store the zip file path in a temporary location accessible by the download endpoint
                # This is a simplification for this environment. A real app would use a shared file store.
                os.rename(zip_filepath, os.path.join('uploads', zip_filename))

                return jsonify(analysis_results)

            except Exception as e:
                return jsonify({"error": f"An unexpected error occurred during analysis: {str(e)}"}), 500

    return jsonify({"error": "An unknown error occurred"}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_report(filename):
    """
    Serves the generated zip file for download.
    """
    # For security, only allow downloads from the 'uploads' directory
    directory = os.path.join(os.getcwd(), 'uploads')
    try:
        return send_from_directory(directory, filename, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found."}), 404

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True) # Ensure uploads dir exists for zips
    app.run(debug=True, port=5001)
