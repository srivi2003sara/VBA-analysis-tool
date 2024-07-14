from flask import Flask, request, jsonify, render_template, redirect, url_for
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsm'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        try:
            vba_code, analysis_results = analyze_vba_code(file_path)  # Perform analysis
            flowchart_path = generate_flowchart(vba_code)  # Generate flowchart
            pdf_filename = generate_pdf_documentation(vba_code, analysis_results)  # Generate PDF

            # Pass the data to the results page
            return jsonify({'success': True, 'redirect': url_for('results')})
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/results')
def results():
    # Replace with actual results and pass them to the template
    vba_code = 'Sample VBA code'
    analysis_results = ['Result 1', 'Result 2', 'Result 3']
    flowchart_path = 'flowchart.png'
    pdf_filename = 'documentation.pdf'
    return render_template('results.html', vba_code=vba_code, analysis_results=analysis_results, flowchart_path=flowchart_path, pdf_filename=pdf_filename, documentation=['Documentation Item'])
