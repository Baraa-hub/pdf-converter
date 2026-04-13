from flask import Flask, request, send_file, render_template
import os, uuid, zipfile
from werkzeug.utils import secure_filename
from pypdf import PdfReader
import pdfplumber

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400
    file = request.files['file']
    fmt = request.form.get('format', 'txt').lower()
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return {'error': 'Please upload a valid PDF'}, 400

    uid = str(uuid.uuid4())[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}.pdf')
    file.save(input_path)

    base_name = secure_filename(file.filename).replace('.pdf', '')
    output_filename = f'{base_name}_converted.{fmt}'
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    try:
        if fmt == 'txt':
            with pdfplumber.open(input_path) as pdf:
                text = '\n\n'.join(p.extract_text() or '' for p in pdf.pages)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text)

        elif fmt == 'html':
            with pdfplumber.open(input_path) as pdf:
                text = '\n'.join(f'<p>{p.extract_text() or ""}</p>' for p in pdf.pages)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(f'<html><body>{text}</body></html>')

        elif fmt == 'md':
            with pdfplumber.open(input_path) as pdf:
                text = '\n\n'.join(p.extract_text() or '' for p in pdf.pages)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text)

        elif fmt == 'csv':
            import csv
            with pdfplumber.open(input_path) as pdf:
                rows = []
                for p in pdf.pages:
                    tables = p.extract_tables()
                    for table in tables:
                        rows.extend(table)
                if not rows:
                    text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
                    rows = [[line] for line in text.split('\n') if line.strip()]
            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(rows)

        elif fmt == 'json':
            import json
            with pdfplumber.open(input_path) as pdf:
                data = [{'page': i+1, 'text': p.extract_text() or ''} for i, p in enumerate(pdf.pages)]
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)

        elif fmt == 'docx':
            from docx import Document
            doc = Document()
            with pdfplumber.open(input_path) as pdf:
                for i, p in enumerate(pdf.pages):
                    doc.add_heading(f'Page {i+1}', level=1)
                    doc.add_paragraph(p.extract_text() or '')
            doc.save(output_path)

        elif fmt in ('jpg', 'png'):
            from pdf2image import convert_from_path
            images = convert_from_path(input_path)
            if len(images) == 1:
                images[0].save(output_path, fmt.upper() if fmt == 'png' else 'JPEG')
            else:
                zip_path = output_path.replace(f'.{fmt}', '.zip')
                output_filename = output_filename.replace(f'.{fmt}', '.zip')
                with zipfile.ZipFile(zip_path, 'w') as zf:
                    for i, img in enumerate(images):
                        img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_page{i+1}.{fmt}')
                        img.save(img_path, fmt.upper() if fmt == 'png' else 'JPEG')
                        zf.write(img_path, f'page{i+1}.{fmt}')
                output_path = zip_path

        else:
            return {'error': 'Unsupported format'}, 400

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        return {'error': str(e)}, 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
