from flask import Flask, request, send_file, render_template
import os, uuid, zipfile
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from PIL import Image

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def pdf_to_images(input_path, dpi=200):
    return convert_from_path(input_path, dpi=dpi)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400
    file = request.files['file']
    fmt = request.form.get('format', 'jpg').lower()
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return {'error': 'Please upload a valid PDF'}, 400

    uid = str(uuid.uuid4())[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}.pdf')
    file.save(input_path)
    base_name = secure_filename(file.filename).replace('.pdf', '')

    try:
        images = pdf_to_images(input_path)

        if fmt in ('jpg', 'png'):
            save_fmt = 'PNG' if fmt == 'png' else 'JPEG'
            if len(images) == 1:
                output_filename = f'{base_name}_converted.{fmt}'
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                images[0].save(output_path, save_fmt)
            else:
                output_filename = f'{base_name}_converted.zip'
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                with zipfile.ZipFile(output_path, 'w') as zf:
                    for i, img in enumerate(images):
                        img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.{fmt}')
                        img.save(img_path, save_fmt)
                        zf.write(img_path, f'page{i+1}.{fmt}')

        elif fmt == 'docx':
            from docx import Document
            from docx.shared import Inches
            output_filename = f'{base_name}_converted.docx'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            doc = Document()
            for i, img in enumerate(images):
                img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.png')
                img.save(img_path, 'PNG')
                section = doc.sections[0] if i == 0 else doc.add_section()
                section.page_width = doc.sections[0].page_width
                section.page_height = doc.sections[0].page_height
                section.top_margin = section.bottom_margin = 914400  # 1 inch
                section.left_margin = section.right_margin = 914400
                if i > 0:
                    doc.add_page_break()
                doc.add_picture(img_path, width=Inches(6.5))
            doc.save(output_path)

        elif fmt == 'pptx':
            from pptx import Presentation
            from pptx.util import Inches, Emu
            output_filename = f'{base_name}_converted.pptx'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            first_img = images[0]
            img_w, img_h = first_img.size
            aspect = img_h / img_w
            slide_w = Inches(10)
            slide_h = Emu(int(slide_w * aspect))
            prs = Presentation()
            prs.slide_width = slide_w
            prs.slide_height = slide_h
            blank_layout = prs.slide_layouts[6]
            for i, img in enumerate(images):
                img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.png')
                img.save(img_path, 'PNG')
                slide = prs.slides.add_slide(blank_layout)
                slide.shapes.add_picture(img_path, 0, 0, width=slide_w, height=slide_h)
            prs.save(output_path)

        elif fmt == 'html':
            import base64
            output_filename = f'{base_name}_converted.html'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            pages_html = ''
            for i, img in enumerate(images):
                img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.png')
                img.save(img_path, 'PNG')
                with open(img_path, 'rb') as f:
                    b64 = base64.b64encode(f.read()).decode()
                pages_html += f'<div class="page"><img src="data:image/png;base64,{b64}" alt="Page {i+1}"></div>\n'
            html = f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
body{{margin:0;padding:20px;background:#666;display:flex;flex-direction:column;align-items:center}}
.page{{margin:10px 0;box-shadow:0 4px 12px rgba(0,0,0,0.4)}}
.page img{{display:block;max-width:900px;width:100%}}
</style></head>
<body>{pages_html}</body></html>'''
            with open(output_path, 'w') as f:
                f.write(html)

        else:
            return {'error': 'Unsupported format'}, 400

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        return {'error': str(e)}, 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
