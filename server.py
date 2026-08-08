"""
Legal Registry Automation Web Server (Flask)
=============================================
Zero-assumption, 100% dynamic deed generation server.
When no diagram is provided, width and length remain strictly blank ("").
"""

import os
import sys
import json
import uuid
import re
import datetime
from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename

from generate_registry import (
    RegistryGenerator, 
    parse_excel_registry_data, 
    parse_plot_dimensions,
    number_to_hindi_words,
    format_indian_currency,
    plot_number_to_hindi,
    clean_plot_no,
    get_today_hindi_date
)
from kyc_parser import parse_kyc_details
from excel_boundary_parser import parse_boundaries_from_excel

app = Flask(__name__, static_folder="static", static_url_path="")
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), "uploads")
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), "outputs")
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

@app.route("/")
def index():
    return send_from_directory("static", "index.html")

@app.route("/api/extract", methods=["POST"])
def api_extract():
    session_id = str(uuid.uuid4())[:8]
    session_dir = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    os.makedirs(session_dir, exist_ok=True)

    excel_file = request.files.get("excel_file")
    png_file = request.files.get("png_file")
    template_file = request.files.get("template_file")
    kyc_files = request.files.getlist("kyc_files")

    excel_path = None
    png_path = None
    template_path = None
    kyc_paths = []

    if excel_file and excel_file.filename:
        excel_path = os.path.join(session_dir, secure_filename(excel_file.filename))
        excel_file.save(excel_path)

    if png_file and png_file.filename:
        png_path = os.path.join(session_dir, secure_filename(png_file.filename))
        png_file.save(png_path)

    if template_file and template_file.filename:
        template_path = os.path.join(session_dir, secure_filename(template_file.filename))
        template_file.save(template_path)

    for kf in kyc_files:
        if kf and kf.filename:
            kp = os.path.join(session_dir, secure_filename(kf.filename))
            kf.save(kp)
            kyc_paths.append(kp)

    today_date_hindi = get_today_hindi_date()

    # If no files uploaded, return pure empty structure
    if not excel_path and not png_path and not kyc_paths and not template_path:
        return jsonify({
            'session_id': session_id,
            'template_path': '',
            'plot_data': {'plot_no': '', 'plot_no_hindi': '', 'width': '', 'length': '', 'area_sqm': '', 'area_sqft': ''},
            'colony_name': '',
            'allotment_val': '',
            'allotment_val_words': '',
            'allotment_val_formatted': '',
            'boundaries': {'east': '', 'west': '', 'north': '', 'south': ''},
            'payments': [],
            'tds_info': None,
            'kyc_data': {'buyers': [{'name': '', 'relation_title': 'पति श्री', 'relation_name': '', 'pan_no': '', 'aadhar_no': '', 'address': ''}], 'primary_name': '', 'relation_title': 'पति श्री', 'relation_name': '', 'pan_no': '', 'aadhar_no': '', 'address': ''},
            'execution_date': today_date_hindi
        })

    # 1. Direct Excel Data Extraction
    excel_data = parse_excel_registry_data(excel_path) if excel_path else {'colony_name': '', 'plot_no': '', 'allotment_val': '', 'payments': [], 'tds_info': None}
    boundaries = parse_boundaries_from_excel(excel_path) if excel_path else {'east': '', 'west': '', 'north': '', 'south': ''}

    # 2. Plot Dimensions (Strict PNG Precedence: If no PNG, width and length remain strictly blank)
    if png_path:
        plot_data = parse_plot_dimensions(png_path)
    elif excel_path:
        raw_p_no = excel_data.get('plot_no', '')
        clean_p = clean_plot_no(raw_p_no)
        p_hindi = plot_number_to_hindi(raw_p_no) if raw_p_no else ''
        plot_data = {
            'plot_no': clean_p or raw_p_no,
            'plot_no_hindi': p_hindi,
            'width': '',
            'length': '',
            'area_sqm': excel_data.get('area_sqm', ''),
            'area_sqft': excel_data.get('area_sqft', '')
        }
    else:
        plot_data = {'plot_no': '', 'plot_no_hindi': '', 'width': '', 'length': '', 'area_sqm': '', 'area_sqft': ''}

    # 3. KYC Extraction (from PDF or Excel)
    kyc_data = parse_kyc_details(kyc_paths, excel_path) if (kyc_paths or excel_path) else {'buyers': [{'name': '', 'relation_title': 'पति श्री', 'relation_name': '', 'pan_no': '', 'aadhar_no': '', 'address': ''}], 'primary_name': '', 'relation_title': 'पति श्री', 'relation_name': '', 'pan_no': '', 'aadhar_no': '', 'address': ''}

    allotment_val = excel_data.get('allotment_val', '')
    response_payload = {
        'session_id': session_id,
        'template_path': template_path or '',
        'plot_data': plot_data,
        'colony_name': excel_data.get('colony_name', ''),
        'allotment_val': allotment_val,
        'allotment_val_words': number_to_hindi_words(allotment_val) if allotment_val else '',
        'allotment_val_formatted': format_indian_currency(allotment_val) if allotment_val else '',
        'boundaries': boundaries,
        'payments': excel_data.get('payments', []),
        'tds_info': excel_data.get('tds_info', None),
        'kyc_data': kyc_data,
        'execution_date': today_date_hindi
    }

    return jsonify(response_payload)

@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.json or {}
    template_path = data.get("template_path")
    if not template_path or not os.path.exists(template_path):
        template_path = os.path.join(os.path.dirname(__file__), "Plot Registry EMPTY.docx")

    raw_p_num = str(data.get('plot_data', {}).get('plot_no', 'Deed'))
    clean_p_num = re.sub(r'[^a-zA-Z0-9_\-]', '_', raw_p_num)
    output_filename = f"Plot_Registry_Filled_{clean_p_num}_{uuid.uuid4().hex[:6]}.docx"
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

    generator = RegistryGenerator(template_path)
    generator.generate(data, output_path)

    return jsonify({
        'status': 'success',
        'filename': output_filename,
        'download_url': f"/api/download/{output_filename}"
    })

@app.route("/api/download/<filename>")
def api_download(filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name=filename)
    return jsonify({'error': 'File not found'}), 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
