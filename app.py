from flask import Flask, request, send_file, render_template, jsonify
import os, uuid, zipfile, base64
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from PIL import Image
import pdfplumber

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def detect_pdf_type(input_path):
    """Returns 'text', 'scanned', or 'mixed'"""
    try:
        with pdfplumber.open(input_path) as pdf:
            has_text = False
            has_images = False
            for page in pdf.pages:
                text = page.extract_text()
                if text and len(text.strip()) > 50:
                    has_text = True
                if page.images:
                    has_images = True
                if has_text and has_images:
                    break
            if has_text and has_images:
                return 'mixed'
            elif has_text:
                return 'text'
            else:
                return 'scanned'
    except:
        return 'scanned'

def pdf_to_images(input_path, dpi=120):
    return convert_from_path(input_path, dpi=dpi, thread_count=1, use_cropbox=True, strict=False)

def ocr_images(images):
    import pytesseract
    return [pytesseract.image_to_string(img, lang='eng+ara') for img in images]

def save_as_docx_images(images, output_path, uid):
    from docx import Document
    from docx.shared import Inches
    doc = Document()
    for i, img in enumerate(images):
        img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.png')
        img.save(img_path, 'PNG')
        if i > 0:
            doc.add_page_break()
        doc.add_picture(img_path, width=Inches(6.5))
    doc.save(output_path)

def save_as_docx_text(input_path, output_path):
    from docx import Document
    from docx.shared import Pt, Inches
    import pdfplumber

    doc = Document()

    with pdfplumber.open(input_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if i > 0:
                doc.add_page_break()

            section = doc.sections[-1]
            page_width_inch = float(page.width) / 72
            page_height_inch = float(page.height) / 72
            section.page_width = Inches(page_width_inch)
            section.page_height = Inches(page_height_inch)
            section.top_margin = Inches(0)
            section.bottom_margin = Inches(0)
            section.left_margin = Inches(0)
            section.right_margin = Inches(0)

            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
                extra_attrs=["fontname", "size"]
            )

            if not words:
                # Scanned page — fall back to OCR
                from pdf2image import convert_from_path
                import pytesseract
                images = convert_from_path(input_path, dpi=150, first_page=i+1, last_page=i+1)
                if images:
                    text = pytesseract.image_to_string(images[0], lang='eng+ara')
                    for line in text.split('\n'):
                        if line.strip():
                            para = doc.add_paragraph()
                            para.paragraph_format.space_before = Pt(0)
                            para.paragraph_format.space_after = Pt(2)
                            run = para.add_run(line)
                            run.font.size = Pt(11)
                continue

            # Group words into lines by y position
            lines = {}
            for word in words:
                y_key = round(float(word['top']) / 5) * 5
                if y_key not in lines:
                    lines[y_key] = []
                lines[y_key].append(word)

            for y_key in sorted(lines.keys()):
                line_words = sorted(lines[y_key], key=lambda w: float(w['x0']))
                para = doc.add_paragraph()
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)

                first_x = float(line_words[0]['x0'])
                left_indent_inch = first_x / 72
                para.paragraph_format.left_indent = Inches(min(left_indent_inch, page_width_inch - 0.5))

                for word in line_words:
                    run = para.add_run(word['text'] + ' ')
                    try:
                        size = float(word.get('size', 12))
                        run.font.size = Pt(max(6, min(size, 72)))
                    except:
                        run.font.size = Pt(12)
                    try:
                        fontname = word.get('fontname', '')
                        if fontname:
                            clean = fontname.split('+')[-1].split('-')[0]
                            run.font.name = clean
                    except:
                        pass

    doc.save(output_path)

def save_as_docx_native(input_path, output_path, uid):
    from docx import Document
    from docx.shared import Inches
    with pdfplumber.open(input_path) as pdf:
        doc = Document()
        for i, page in enumerate(pdf.pages):
            if i > 0:
                doc.add_page_break()
            text = page.extract_text() or ''
            for line in text.split('\n'):
                if line.strip():
                    doc.add_paragraph(line)
    doc.save(output_path)

def save_as_pptx_images(images, output_path, uid):
    from pptx import Presentation
    from pptx.util import Inches, Emu
    first_img = images[0]
    img_w, img_h = first_img.size
    slide_w = Inches(10)
    slide_h = Emu(int(slide_w * img_h / img_w))
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

def save_as_html_images(images, output_path, uid):
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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/detect', methods=['POST'])
def detect():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    if not file.filename.endswith('.pdf'):
        return jsonify({'error': 'Not a PDF'}), 400
    uid = str(uuid.uuid4())[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}.pdf')
    file.save(input_path)
    pdf_type = detect_pdf_type(input_path)
    return jsonify({'type': pdf_type, 'uid': uid})

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400
    file = request.files['file']
    fmt = request.form.get('format', 'jpg').lower()
    mode = request.form.get('mode', 'image')
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return {'error': 'Please upload a valid PDF'}, 400

    uid = str(uuid.uuid4())[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}.pdf')
    file.save(input_path)
    base_name = secure_filename(file.filename).replace('.pdf', '')

    try:
        page_count = 0
        with pdfplumber.open(input_path) as pdf:
            page_count = len(pdf.pages)
        dpi = 150 if page_count <= 5 else 80
        images = pdf_to_images(input_path, dpi=dpi)

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
            output_filename = f'{base_name}_converted.docx'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            if mode == 'ocr':
                pages_text = ocr_images(images)
                save_as_docx_text(input_path, output_path)
            elif mode == 'native':
                save_as_docx_native(input_path, output_path, uid)
            else:
                save_as_docx_images(images, output_path, uid)

        elif fmt == 'pptx':
            output_filename = f'{base_name}_converted.pptx'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            if mode == 'ocr':
                from pptx import Presentation
                from pptx.util import Inches, Emu
                pages_text = ocr_images(images)
                first_img = images[0]
                slide_w = Inches(10)
                slide_h = Emu(int(slide_w * first_img.size[1] / first_img.size[0]))
                prs = Presentation()
                prs.slide_width = slide_w
                prs.slide_height = slide_h
                slide_layout = prs.slide_layouts[1]
                for i, text in enumerate(pages_text):
                    slide = prs.slides.add_slide(slide_layout)
                    slide.shapes.title.text = f'Page {i+1}'
                    tf = slide.placeholders[1].text_frame
                    tf.word_wrap = True
                    tf.text = text[:800] if text.strip() else '(no text detected)'
                prs.save(output_path)
            else:
                save_as_pptx_images(images, output_path, uid)

        elif fmt == 'html':
            output_filename = f'{base_name}_converted.html'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            save_as_html_images(images, output_path, uid)

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
