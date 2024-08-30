from flask import Flask, request, render_template, send_file, redirect, url_for, flash
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx2pdf import convert
import os

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key_here'  # Set a secret key for session management

# Directories for file handling
UPLOAD_FOLDER = 'uploads/'
CERTIFICATES_FOLDER = 'certificates/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CERTIFICATES_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if the uploaded file is allowed."""
    return filename.lower().endswith(('.xlsx', '.docx'))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'excel_file' not in request.files or 'template_file' not in request.files:
        flash('Files are required', 'error')
        return redirect(url_for('index'))
    
    excel_file = request.files['excel_file']
    template_file = request.files['template_file']
    output_format = request.form.get('output_format')
    
    if not allowed_file(excel_file.filename) or not allowed_file(template_file.filename):
        flash('Invalid file type', 'error')
        return redirect(url_for('index'))

    excel_path = os.path.join(UPLOAD_FOLDER, 'data.xlsx')
    template_path = os.path.join(UPLOAD_FOLDER, 'template.docx')
    
    excel_file.save(excel_path)
    template_file.save(template_path)
    
    df = pd.read_excel(excel_path)
    
    for index, row in df.iterrows():
        doc = Document(template_path)
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                # Example: replace {{ Name }} with actual name
                if '{{ Name }}' in run.text:
                    run.text = run.text.replace('{{ Name }}', str(row.get('Name', '')))
                    run.bold = True
                    run.italic = True
                    run.font.size = Pt(16)
        
        docx_path = os.path.join(CERTIFICATES_FOLDER, f'{row.get("Name", "Certificate")}_certificate.docx')
        if output_format == 'PDF':
            pdf_path = os.path.join(CERTIFICATES_FOLDER, f'{row.get("Name", "Certificate")}_certificate.pdf')
            doc.save(docx_path)
            convert(docx_path, pdf_path)
            os.remove(docx_path)
            return send_file(pdf_path, as_attachment=True)
        else:
            doc.save(docx_path)
            return send_file(docx_path, as_attachment=True)
    
    flash('Certificates generated successfully!', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
    port = int(os.environ.get('PORT', 5000))  # Get the PORT environment variable, default to 5000
    app.run(host='0.0.0.0', port=port)
