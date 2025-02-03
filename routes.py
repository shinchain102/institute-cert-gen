import os
import logging
from flask import render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from app import app
from utils import process_template, process_data_file, generate_certificates

ALLOWED_TEMPLATE_EXTENSIONS = {'docx', 'ppt', 'pptx'}
ALLOWED_DATA_EXTENSIONS = {'csv', 'xlsx'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload-template', methods=['POST'])
def upload_template():
    if 'template' not in request.files:
        return jsonify({'error': 'No template file provided'}), 400

    file = request.files['template']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename, ALLOWED_TEMPLATE_EXTENSIONS):
        return jsonify({'error': 'Invalid file type for template'}), 400

    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        app.logger.info(f"Template file saved at: {filepath}")

        variables = process_template(filepath)
        return jsonify({
            'message': 'Template uploaded successfully',
            'variables': variables,
            'filename': filename  # Send filename back to client
        })
    except Exception as e:
        app.logger.error(f'Template upload error: {str(e)}')
        return jsonify({'error': str(e)}), 500

@app.route('/upload-data', methods=['POST'])
def upload_data():
    if 'data' not in request.files:
        return jsonify({'error': 'No data file provided'}), 400

    file = request.files['data']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename, ALLOWED_DATA_EXTENSIONS):
        return jsonify({'error': 'Invalid file type for data file'}), 400

    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        app.logger.info(f"Data file saved at: {filepath}")

        columns = process_data_file(filepath)
        return jsonify({
            'message': 'Data file uploaded successfully',
            'columns': columns,
            'filename': filename  # Send filename back to client
        })
    except Exception as e:
        app.logger.error(f'Data file upload error: {str(e)}')
        return jsonify({'error': str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        if not data:
            return jsonify({'error': 'No data provided'}), 400

        mapping = data.get('mapping')
        template_file = data.get('template_file')
        data_file = data.get('data_file')

        if not all([mapping, template_file, data_file]):
            return jsonify({'error': 'Missing required parameters'}), 400

        # Construct full file paths
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_file)
        data_path = os.path.join(app.config['UPLOAD_FOLDER'], data_file)

        # Verify files exist
        if not os.path.exists(template_path):
            return jsonify({'error': f'Template file not found: {template_file}'}), 404
        if not os.path.exists(data_path):
            return jsonify({'error': f'Data file not found: {data_file}'}), 404

        app.logger.info(f"Generating certificates using template: {template_path} and data: {data_path}")
        app.logger.info(f"Variable mapping: {mapping}")

        try:
            zip_path = generate_certificates(template_path, data_path, mapping)
            return send_file(zip_path, as_attachment=True, download_name='certificates.zip')
        except Exception as gen_error:
            app.logger.error(f"Certificate generation error: {str(gen_error)}")
            return jsonify({'error': str(gen_error)}), 500

    except Exception as e:
        app.logger.error(f"Request processing error: {str(e)}")
        return jsonify({'error': 'Failed to process the request'}), 500