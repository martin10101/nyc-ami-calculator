import os
from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename
from main import main as run_ami_optix_analysis
from ami_optix.narrator import generate_narrative

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # 16 MB max upload size

# Ensure the upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/analyze', methods=['POST'])
def analyze_file():
    """
    API endpoint to upload a project spreadsheet and get the optimization analysis.
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "No file selected for uploading"}), 400

    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        try:
            # Get LLM parameters from the form data
            provider = request.form.get('provider', 'openai') # default to openai
            model_name = request.form.get('model_name', 'gpt-4') # default to gpt-4

            # Call the core analysis function from main.py
            analysis_dict = run_ami_optix_analysis(filepath)

            if "error" in analysis_dict:
                return jsonify(analysis_dict), 400

            # Call the narrator to get the analysis text
            narrative = generate_narrative(analysis_dict, provider, model_name)

            # Add the narrative to the final result
            analysis_dict['narrative_analysis'] = narrative

            return jsonify(analysis_dict)
        except Exception as e:
            return jsonify({"error": f"An unexpected error occurred during analysis: {str(e)}"}), 500
        finally:
            # Clean up the uploaded file after analysis
            if os.path.exists(filepath):
                os.remove(filepath)

    return jsonify({"error": "An unknown error occurred"}), 500

if __name__ == '__main__':
    # Note: This is a development server. For production, use a proper WSGI server like Gunicorn.
    app.run(debug=True, port=5001)
