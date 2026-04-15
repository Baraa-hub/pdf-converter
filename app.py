from flask import Flask, request, send_file, render_template, jsonify
import os, uuid, zipfile, base64
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from PIL import Image
import pdfplumber

app = Flask(__name__, static_folder='static')
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def detect_pdf_type(input_path):
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

def is_rtl_text(text):
    import unicodedata
    return any(unicodedata.bidirectional(c) in ('R', 'AL') for c in text if c.strip())

def fix_rtl(line):
    from bidi.algorithm import get_display
    return get_display(line)

def save_image_file(img, path, save_fmt):
    img = img.copy()
    if save_fmt == 'JPEG':
        if img.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            if img.mode in ('RGBA', 'LA'):
                background.paste(img, mask=img.split()[-1])
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        img.save(path, 'JPEG', quality=95)
    else:
        if img.mode not in ('RGB', 'RGBA'):
            img = img.convert('RGB')
        img.save(path, 'PNG')

def save_as_docx_images(images, output_path, uid):
    from docx import Document
    from docx.shared import Inches
    doc = Document()
    for i, img in enumerate(images):
        img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.png')
        save_image_file(img, img_path, 'PNG')
        if i > 0:
            doc.add_page_break()
        doc.add_picture(img_path, width=Inches(6.5))
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
        save_image_file(img, img_path, 'PNG')
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(img_path, 0, 0, width=slide_w, height=slide_h)
    prs.save(output_path)

def save_as_html_images(images, output_path, uid):
    pages_html = ''
    for i, img in enumerate(images):
        img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.png')
        save_image_file(img, img_path, 'PNG')
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

# ── Helpers for native DOCX conversion ────────────────────────────────────────

def get_font_size(word):
    try:
        return float(word.get('size', 12))
    except:
        return 12.0

def is_bold(word):
    fontname = word.get('fontname', '')
    return any(b in fontname.lower() for b in ['bold', 'black', 'heavy', 'semibold', 'demi'])

def apply_rtl_to_paragraph(para):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    pPr = para._p.get_or_add_pPr()
    bidi_el = OxmlElement('w:bidi')
    pPr.append(bidi_el)
    para.alignment = 2  # right align

def apply_rtl_to_run(run):
    from docx.oxml import OxmlElement
    rPr = run._r.get_or_add_rPr()
    rtl_el = OxmlElement('w:rtl')
    rPr.append(rtl_el)

def set_cell_background(cell, hex_color):
    """Set background color of a table cell."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_run_to_cell(cell, text, bold=False, font_size=11, rtl=False):
    from docx.shared import Pt
    para = cell.paragraphs[0]
    if rtl:
        apply_rtl_to_paragraph(para)
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(font_size)
    if rtl:
        apply_rtl_to_run(run)

def compute_median_font_size(page):
    """Get the median font size on a page to detect headings."""
    sizes = []
    try:
        words = page.extract_words(extra_attrs=['size'])
        for w in words:
            try:
                sizes.append(float(w.get('size', 12)))
            except:
                pass
    except:
        pass
    if not sizes:
        return 12.0
    sizes.sort()
    return sizes[len(sizes) // 2]

def bbox_overlaps_table(bbox, table_bboxes):
    """Check if a bounding box overlaps with any table bounding box."""
    x0, top, x1, bottom = bbox
    for tb in table_bboxes:
        tx0, ttop, tx1, tbottom = tb
        if x0 < tx1 and x1 > tx0 and top < tbottom and bottom > ttop:
            return True
    return False

def extract_page_images(input_path, page_index, uid, output_folder):
    """Extract embedded images from a PDF page and save them to disk."""
    import fitz  # PyMuPDF
    saved = []
    try:
        doc = fitz.open(input_path)
        page = doc[page_index]
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            img_bytes = base_image['image']
            img_ext = base_image['ext']
            img_path = os.path.join(output_folder, f'{uid}_page{page_index}_img{img_index}.{img_ext}')
            with open(img_path, 'wb') as f:
                f.write(img_bytes)
            saved.append(img_path)
        doc.close()
    except Exception:
        pass
    return saved

def save_as_docx_native(input_path, output_path, uid):
    """
    Generic native PDF-to-DOCX conversion.
    - Extracts real tables and rebuilds them as DOCX tables
    - Preserves bold, font sizes, headings
    - Handles RTL Arabic text
    - Embeds images extracted from the PDF
    - Falls back to OCR for scanned pages
    """
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.oxml.ns import qn

    # Try importing PyMuPDF for image extraction — optional
    try:
        import fitz
        has_fitz = True
    except ImportError:
        has_fitz = False

    doc = Document()

    # Set default margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    with pdfplumber.open(input_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            if page_index > 0:
                doc.add_page_break()

            # Get median font size for heading detection
            median_size = compute_median_font_size(page)

            # Extract tables on this page
            tables = page.extract_tables({
                'vertical_strategy': 'lines',
                'horizontal_strategy': 'lines',
                'snap_tolerance': 3,
                'join_tolerance': 3,
                'edge_min_length': 3,
                'min_words_vertical': 1,
                'min_words_horizontal': 1,
            })

            # Get table bounding boxes to exclude from text extraction
            table_bboxes = []
            try:
                for table_obj in page.find_tables():
                    table_bboxes.append(table_obj.bbox)
            except:
                pass

            # Extract words outside tables
            try:
                all_words = page.extract_words(
                    x_tolerance=3,
                    y_tolerance=3,
                    keep_blank_chars=False,
                    use_text_flow=False,
                    extra_attrs=['fontname', 'size']
                )
            except:
                all_words = []

            # Filter out words that fall inside table bboxes
            text_words = [
                w for w in all_words
                if not bbox_overlaps_table(
                    (float(w['x0']), float(w['top']), float(w['x1']), float(w['bottom'])),
                    table_bboxes
                )
            ]

            # If no native text at all, fall back to OCR
            if not all_words:
                try:
                    import pytesseract
                    imgs = convert_from_path(input_path, dpi=150, first_page=page_index+1, last_page=page_index+1)
                    if imgs:
                        text = pytesseract.image_to_string(imgs[0], lang='eng+ara')
                        for line in text.split('\n'):
                            line = line.strip()
                            if line:
                                para = doc.add_paragraph()
                                rtl = is_rtl_text(line)
                                if rtl:
                                    line = fix_rtl(line)
                                    apply_rtl_to_paragraph(para)
                                run = para.add_run(line)
                                run.font.size = Pt(11)
                                if rtl:
                                    apply_rtl_to_run(run)
                except:
                    pass
                continue

            # Embed page images if PyMuPDF available
            if has_fitz:
                img_paths = extract_page_images(input_path, page_index, uid, OUTPUT_FOLDER)
                for img_path in img_paths:
                    try:
                        # Skip tiny images (icons, artifacts)
                        pil_img = Image.open(img_path)
                        w_px, h_px = pil_img.size
                        if w_px < 50 or h_px < 50:
                            continue
                        # Scale to fit page width
                        max_width = Inches(6.5)
                        doc.add_picture(img_path, width=max_width)
                    except:
                        pass

            # Group text words into lines by vertical position
            lines_dict = {}
            for word in text_words:
                y_key = round(float(word['top']) / 4) * 4
                if y_key not in lines_dict:
                    lines_dict[y_key] = []
                lines_dict[y_key].append(word)

            # Track which y positions are covered by tables so we
            # interleave text and tables in reading order
            # Build a list of (y_position, type, content) events
            events = []

            # Add text lines as events
            for y_key, words in lines_dict.items():
                events.append((y_key, 'text', words))

            # Add tables as events using their top y position
            try:
                page_table_objs = list(page.find_tables())
            except:
                page_table_objs = []

            for t_idx, table_obj in enumerate(page_table_objs):
                t_top = table_obj.bbox[1]
                table_data = tables[t_idx] if t_idx < len(tables) else None
                if table_data:
                    events.append((t_top, 'table', table_data))

            # Sort all events by vertical position
            events.sort(key=lambda e: e[0])

            for _, event_type, content in events:
                if event_type == 'text':
                    words = sorted(content, key=lambda w: float(w['x0']))
                    if not words:
                        continue

                    line_text = ' '.join(w['text'] for w in words).strip()
                    if not line_text:
                        continue

                    # Detect font properties from the majority of words
                    sizes = [get_font_size(w) for w in words]
                    avg_size = sum(sizes) / len(sizes) if sizes else 12.0
                    bold_count = sum(1 for w in words if is_bold(w))
                    mostly_bold = bold_count > len(words) / 2
                    rtl = is_rtl_text(line_text)

                    if rtl:
                        line_text = fix_rtl(line_text)

                    para = doc.add_paragraph()
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(2)

                    if rtl:
                        apply_rtl_to_paragraph(para)

                    run = para.add_run(line_text)
                    run.bold = mostly_bold

                    # Heading detection: significantly larger than median
                    if avg_size > median_size * 1.3:
                        run.font.size = Pt(min(avg_size, 24))
                        run.bold = True
                    else:
                        run.font.size = Pt(max(8, min(avg_size, 72)))

                    if rtl:
                        apply_rtl_to_run(run)

                elif event_type == 'table':
                    table_data = content
                    if not table_data:
                        continue

                    # Filter completely empty rows
                    rows = [r for r in table_data if any(c and str(c).strip() for c in r)]
                    if not rows:
                        continue

                    num_cols = max(len(r) for r in rows)
                    if num_cols == 0:
                        continue

                    # Normalize rows to same column count
                    norm_rows = []
                    for row in rows:
                        norm_row = list(row) + [None] * (num_cols - len(row))
                        norm_rows.append(norm_row)

                    tbl = doc.add_table(rows=len(norm_rows), cols=num_cols)
                    tbl.style = 'Table Grid'

                    for r_idx, row in enumerate(norm_rows):
                        for c_idx, cell_val in enumerate(row):
                            cell = tbl.rows[r_idx].cells[c_idx]
                            text = str(cell_val).strip() if cell_val else ''

                            if not text or text == 'None':
                                continue

                            rtl = is_rtl_text(text)
                            if rtl:
                                text = fix_rtl(text)

                            # Clear default empty paragraph and set text
                            para = cell.paragraphs[0]
                            if rtl:
                                apply_rtl_to_paragraph(para)
                            run = para.add_run(text)
                            run.font.size = Pt(10)
                            if rtl:
                                apply_rtl_to_run(run)

                    doc.add_paragraph()  # spacing after table

    doc.save(output_path)


def save_as_docx_text(input_path, output_path):
    """OCR-based DOCX conversion for scanned PDFs."""
    from docx import Document
    from docx.shared import Pt, Inches

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
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)

            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
                extra_attrs=['fontname', 'size']
            )

            if not words:
                import pytesseract
                imgs = convert_from_path(input_path, dpi=150, first_page=i+1, last_page=i+1)
                if imgs:
                    text = pytesseract.image_to_string(imgs[0], lang='eng+ara')
                    for line in text.split('\n'):
                        if line.strip():
                            para = doc.add_paragraph()
                            rtl = is_rtl_text(line)
                            if rtl:
                                line = fix_rtl(line)
                                apply_rtl_to_paragraph(para)
                            run = para.add_run(line)
                            run.font.size = Pt(11)
                            if rtl:
                                apply_rtl_to_run(run)
                continue

            native_text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if native_text and native_text.strip():
                for line in native_text.split('\n'):
                    if line.strip():
                        para = doc.add_paragraph()
                        rtl = is_rtl_text(line)
                        if rtl:
                            line = fix_rtl(line)
                            apply_rtl_to_paragraph(para)
                        run = para.add_run(line)
                        run.font.size = Pt(12)
                        if rtl:
                            apply_rtl_to_run(run)
                continue

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
                for word in line_words:
                    run = para.add_run(word['text'] + ' ')
                    try:
                        size = float(word.get('size', 12))
                        run.font.size = Pt(max(6, min(size, 72)))
                    except:
                        run.font.size = Pt(12)

    doc.save(output_path)


def parse_pages(pages_param, total):
    indices = []
    for part in pages_param.split(','):
        part = part.strip()
        if '-' in part:
            try:
                parts = part.split('-')
                start = int(parts[0].strip())
                end = int(parts[1].strip())
                indices += list(range(start - 1, end))
            except:
                pass
        else:
            try:
                indices.append(int(part) - 1)
            except:
                pass
    return sorted(set([i for i in indices if 0 <= i < total]))

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
    with pdfplumber.open(input_path) as pdf:
        page_count = len(pdf.pages)
    return jsonify({'type': pdf_type, 'uid': uid, 'page_count': page_count})

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    fmt = request.form.get('format', 'jpg').lower()
    mode = request.form.get('mode', 'image')
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return jsonify({'error': 'Please upload a valid PDF'}), 400

    uid = str(uuid.uuid4())[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}.pdf')
    file.save(input_path)
    base_name = secure_filename(file.filename).replace('.pdf', '').replace('.', '_')

    try:
        with pdfplumber.open(input_path) as pdf:
            page_count = len(pdf.pages)
        dpi = 150 if page_count <= 5 else 80
        images = pdf_to_images(input_path, dpi=dpi)

        if fmt in ('jpg', 'png'):
            save_fmt = 'JPEG' if fmt == 'jpg' else 'PNG'
            ext = fmt

            pages_param = request.form.get('pages', '').strip()
            if pages_param:
                selected_indices = parse_pages(pages_param, len(images))
                if not selected_indices:
                    return jsonify({'error': 'No valid pages selected. Use format like: 1, 3, 5-7'}), 400
                selected_images = [images[i] for i in selected_indices]
            else:
                selected_images = images

            if len(selected_images) == 1:
                output_filename = f'{base_name}_converted.{ext}'
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                save_image_file(selected_images[0], output_path, save_fmt)
            else:
                output_filename = f'{base_name}_converted.zip'
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                img_paths = []
                for i, img in enumerate(selected_images):
                    img_path = os.path.join(OUTPUT_FOLDER, f'{uid}_p{i+1}.{ext}')
                    save_image_file(img, img_path, save_fmt)
                    img_paths.append((img_path, f'page{i+1}.{ext}'))
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for img_path, arcname in img_paths:
                        zf.write(img_path, arcname)

        elif fmt == 'docx':
            output_filename = f'{base_name}_converted.docx'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            if mode == 'ocr':
                save_as_docx_text(input_path, output_path)
            else:
                # Both 'native' and 'image' modes use the new native converter
                # Image-based fallback only if native fails
                try:
                    save_as_docx_native(input_path, output_path, uid)
                except Exception as e:
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
            return jsonify({'error': 'Unsupported format'}), 400

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
