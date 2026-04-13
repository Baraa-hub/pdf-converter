from flask import Flask, request, send_file, render_template
import os, uuid
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
pytesseract.pytesseract.tesseract_cmd = '/run/current-system/sw/bin/tesseract'

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def pdf_to_images(input_path):
    return convert_from_path(input_path, dpi=200)

def extract_text_ocr(images):
    full_text = []
    for img in images:
        text = pytesseract.image_to_string(img, lang='eng+ara')
        full_text.append(text)
    return full_text

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
        if fmt in ('jpg', 'png'):
            images = pdf_to_images(input_path)
            if len(images) == 1:
                save_fmt = 'PNG' if fmt == 'png' else 'JPEG'
                images[0].save(output_path, save_fmt)
            else:
                import zipfile
                zip_filename = f'{base_name}_converted.zip'
                zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
                with zipfile.ZipFile(zip_path, 'w') as zf:
                    for idx, img in enumerate(images):
                        save_fmt = 'PNG' if fmt == 'png' else 'JPEG'
                        img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{idx+1}.{fmt}')
                        img.save(img_path, save_fmt)
                        zf.write(img_path, f'page{idx+1}.{fmt}')
                output_path = zip_path
                output_filename = zip_filename

        elif fmt == 'html':
            images = pdf_to_images(input_path)
            pages_text = extract_text_ocr(images)
            pages_html = ''
            for i, text in enumerate(pages_text):
                text_html = text.replace('\n', '<br>')
                pages_html += f'<div class="page"><h2>Page {i+1}</h2><p>{text_html}</p></div>'
            html_content = f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
body{{font-family:Arial,sans-serif;max-width:800px;margin:40px auto;padding:20px}}
.page{{margin-bottom:40px;padding:20px;border:1px solid #ddd;border-radius:8px}}
h2{{color:#555;font-size:14px}}
</style></head>
<body>{pages_html}</body></html>'''
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

        elif fmt == 'docx':
            from docx import Document
            images = pdf_to_images(input_path)
            pages_text = extract_text_ocr(images)
            doc = Document()
            doc.add_heading('Converted Document', 0)
            for i, text in enumerate(pages_text):
                doc.add_heading(f'Page {i+1}', level=1)
                for line in text.split('\n'):
                    if line.strip():
                        doc.add_paragraph(line)
                doc.add_page_break()
            doc.save(output_path)

        elif fmt == 'pptx':
            from pptx import Presentation
            from pptx.util import Pt
            images = pdf_to_images(input_path)
            pages_text = extract_text_ocr(images)
            prs = Presentation()
            slide_layout = prs.slide_layouts[1]
            for i, text in enumerate(pages_text):
                slide = prs.slides.add_slide(slide_layout)
                title = slide.shapes.title
                body = slide.placeholders[1]
                title.text = f'Page {i+1}'
                tf = body.text_frame
                tf.word_wrap = True
                tf.text = text[:800] if text.strip() else '(no text detected)'
            output_filename = f'{base_name}_converted.pptx'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            prs.save(output_path)

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
