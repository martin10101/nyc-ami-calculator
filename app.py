import os
import shutil
import tempfile
import zipfile
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from main import main as run_ami_optix_analysis
from ami_optix.narrator import generate_internal_summary
from ami_optix.report_generator import create_excel_reports

app = Flask(__name__)

@app.route('/')
def home():
    return '''
    <html>
    <head>
        <title>NYC AMI Calculator</title>
        <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 50px auto; padding: 20px; }
            .header { text-align: center; color: #2c3e50; }
            .api-info { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; }
            .endpoint { background: #e9ecef; padding: 10px; margin: 10px 0; border-radius: 4px; font-family: monospace; }
        </style>
    </head>
    <body>
        <h1 class="header">🏙️ NYC AMI Calculator</h1>
        <p>Welcome to the NYC Area Median Income (AMI) Calculator API. This service helps analyze affordable housing projects in New York City.</p>
        
        <div class="api-info">
            <h3>Upload File for Analysis:</h3>
            <form id="uploadForm" enctype="multipart/form-data">
                <input type="file" id="fileInput" accept=".csv,.xlsx,.xls" required style="margin: 10px 0; padding: 10px; width: 100%; border: 2px dashed #ccc; border-radius: 4px;">
                <br>
                <button type="submit" style="background: #007bff; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px;">Analyze File</button>
            </form>
            <div id="results" style="margin-top: 20px; display: none;"></div>
        </div>
        
        <div class="api-info">
            <h3>API Endpoints:</h3>
            <div class="endpoint">
                <strong>POST /api/analyze</strong><br>
                Upload a CSV or Excel file for AMI analysis. Returns optimized scenarios and compliance reports.
            </div>
            <div class="endpoint">
                <strong>GET /api/download/&lt;filename&gt;</strong><br>
                Download generated Excel reports as a ZIP file.
            </div>
        </div>
        
        <p><strong>Status:</strong> ✅ Service is running and ready to process files!</p>
        
        <script>
            document.getElementById('uploadForm').addEventListener('submit', async function(e) {
                e.preventDefault();
                const fileInput = document.getElementById('fileInput');
                const resultsDiv = document.getElementById('results');
                
                if (!fileInput.files[0]) {
                    alert('Please select a file');
                    return;
                }
                
                const formData = new FormData();
                formData.append('file', fileInput.files[0]);
                
                resultsDiv.innerHTML = '<p>⏳ Analyzing file... This may take a few minutes.</p>';
                resultsDiv.style.display = 'block';
                
                try {
                    const response = await fetch('/api/analyze', {
                        method: 'POST',
                        body: formData
                    });
                    
                    const data = await response.json();
                    
                    if (response.ok) {
                        resultsDiv.innerHTML = `
                            <h4>✅ Analysis Complete!</h4>
                            <p><strong>Total Units:</strong> ${data.project_summary.total_affordable_units}</p>
                            <p><strong>Total SF:</strong> ${data.project_summary.total_affordable_sf.toLocaleString()} sq ft</p>
                            <p><strong>WAAMI:</strong> ${data.scenario_absolute_best.waami.toFixed(2)}%</p>
                            <p><strong>Bands:</strong> ${data.scenario_absolute_best.bands.join(', ')}</p>
                            ${data.download_link ? `<p><a href="${data.download_link}" style="color: #007bff;">📥 Download Excel Reports</a></p>` : ''}
                            <details style="margin-top: 10px;">
                                <summary>View Full Results</summary>
                                <pre style="background: #f8f9fa; padding: 10px; border-radius: 4px; overflow-x: auto;">${JSON.stringify(data, null, 2)}</pre>
                            </details>
                        `;
                    } else {
                        resultsDiv.innerHTML = `<p style="color: red;">❌ Error: ${data.error}</p>`;
                    }
                } catch (error) {
                    resultsDiv.innerHTML = `<p style="color: red;">❌ Error: ${error.message}</p>`;
                }
            });
        </script>
    </body>
    </html>
    '''

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
                analysis_dict = run_ami_optix_analysis(upload_filepath)
                if "error" in analysis_dict:
                    return jsonify(analysis_dict), 400

                # 2. Generate the internal summary (no LLM for now)
                narrative = generate_internal_summary(analysis_dict)
                analysis_dict['narrative_analysis'] = narrative

                # 3. Generate Excel reports
                report_files = create_excel_reports(analysis_dict, upload_filepath, output_dir=temp_dir)

                # 4. Create a zip file containing all reports
                zip_filename = f"{os.path.splitext(filename)[0]}_reports.zip"
                zip_filepath = os.path.join(temp_dir, zip_filename)
                with zipfile.ZipFile(zip_filepath, 'w') as zipf:
                    for report_file in report_files:
                        zipf.write(report_file, os.path.basename(report_file))

                # 5. Add a download link for the zip file to the response
                analysis_dict['download_link'] = f"/api/download/{zip_filename}"

                # Store the zip file path in a temporary location accessible by the download endpoint
                # This is a simplification for this environment. A real app would use a shared file store.
                uploads_dir = 'uploads'
                os.makedirs(uploads_dir, exist_ok=True)
                shutil.move(zip_filepath, os.path.join(uploads_dir, zip_filename))

                return jsonify(analysis_dict)

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
