from flask import Flask, request, send_file, render_template, jsonify
import os, uuid, zipfile, base64, re, unicodedata
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
    return any(unicodedata.bidirectional(c) in ('R', 'AL') for c in text if c.strip())

def fix_rtl(line):
    from bidi.algorithm import get_display
    return get_display(line)

def clean_text(text):
    if not text:
        return ''
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    text = text.replace('\n', ' ').replace('\r', ' ')
    text = re.sub(r' +', ' ', text).strip()
    return text

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

# ── Helpers ────────────────────────────────────────────────────────────────────

def rgb_to_hex(color):
    try:
        if isinstance(color, (list, tuple)) and len(color) == 3:
            r, g, b = [int(round(c * 255)) for c in color]
            return f'{r:02X}{g:02X}{b:02X}'
        if isinstance(color, (list, tuple)) and len(color) == 1:
            v = int(round(color[0] * 255))
            return f'{v:02X}{v:02X}{v:02X}'
    except:
        pass
    return None

def is_bold(word):
    fontname = word.get('fontname', '')
    return any(b in fontname.lower() for b in ['bold', 'black', 'heavy', 'semibold', 'demi'])

def apply_rtl_to_paragraph(para):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    pPr = para._p.get_or_add_pPr()
    bidi_el = OxmlElement('w:bidi')
    pPr.append(bidi_el)
    para.alignment = 2

def apply_rtl_to_run(run):
    from docx.oxml import OxmlElement
    rPr = run._r.get_or_add_rPr()
    rtl_el = OxmlElement('w:rtl')
    rPr.append(rtl_el)

def set_cell_background(cell, hex_color):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def compute_median_font_size(page):
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

def get_rect_color_at(rects, y_top, y_bottom, x0, x1, page_w, page_h):
    best_color = None
    best_area = 0
    for r in rects:
        if r['width'] > page_w * 0.95 and r['height'] > page_h * 0.5:
            continue
        oy0 = max(r['top'], y_top)
        oy1 = min(r['bottom'], y_bottom)
        ox0 = max(r['x0'], x0)
        ox1 = min(r['x1'], x1)
        if oy1 > oy0 and ox1 > ox0:
            area = (oy1 - oy0) * (ox1 - ox0)
            if area > best_area:
                color = r.get('non_stroking_color')
                if color is not None:
                    hex_c = rgb_to_hex(color)
                    if hex_c and hex_c.upper() not in ('FFFFFF', 'FEFEFE', 'FDFDFD'):
                        best_color = hex_c
                        best_area = area
    return best_color

def merge_split_cells(row):
    """
    Merge cells where text appears to be split across boundaries.
    e.g. ['Contra', 'ct Data', '', ...] → ['Contract Data', '', '', ...]
    Strategy: if cell text doesn't end with space/punctuation and next cell
    starts with lowercase or continues a word, merge them.
    """
    if not row:
        return row

    merged = list(row)
    i = 0
    while i < len(merged) - 1:
        cur = clean_text(str(merged[i]) if merged[i] else '')
        nxt = clean_text(str(merged[i+1]) if merged[i+1] else '')

        if not cur or not nxt:
            i += 1
            continue

        # Detect split: current ends mid-word (no space, no punctuation)
        # and next starts with lowercase or continues naturally
        cur_ends_midword = (
            len(cur) > 0 and
            cur[-1].isalpha() and
            not cur[-1].isspace()
        )
        nxt_starts_midword = (
            len(nxt) > 0 and
            (nxt[0].islower() or nxt[0].isdigit())
        )

        if cur_ends_midword and nxt_starts_midword:
            merged[i] = cur + nxt
            merged.pop(i + 1)
            # Don't increment i — check again in case of triple split
        else:
            i += 1

    return merged

def extract_page_images_pymupdf(input_path, page_index, uid, output_folder):
    """Extract embedded images from a PDF page using PyMuPDF."""
    saved = []
    try:
        import fitz
        doc = fitz.open(input_path)
        page = doc[page_index]
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            try:
                base_image = doc.extract_image(xref)
                img_bytes = base_image['image']
                img_ext = base_image.get('ext', 'png')
                # Skip very small images (icons/artifacts)
                w = base_image.get('width', 0)
                h = base_image.get('height', 0)
                if w < 50 or h < 50:
                    continue
                img_path = os.path.join(output_folder, f'{uid}_pg{page_index}_img{img_index}.{img_ext}')
                with open(img_path, 'wb') as f:
                    f.write(img_bytes)
                saved.append(img_path)
            except:
                pass
        doc.close()
    except ImportError:
        # PyMuPDF not available — skip image extraction
        pass
    except Exception:
        pass
    return saved

# ── Main native DOCX converter ─────────────────────────────────────────────────

def save_as_docx_native(input_path, output_path, uid):
    from docx import Document
    from docx.shared import Pt, Inches

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    with pdfplumber.open(input_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            if page_index > 0:
                doc.add_page_break()

            page_w = float(page.width)
            page_h = float(page.height)
            rects = page.rects
            median_size = compute_median_font_size(page)

            # ── Embed page images before table ─────────────────────────────
            img_paths = extract_page_images_pymupdf(input_path, page_index, uid, OUTPUT_FOLDER)
            for img_path in img_paths:
                try:
                    pil_img = Image.open(img_path)
                    w_px, h_px = pil_img.size
                    if w_px < 50 or h_px < 50:
                        continue
                    max_width = Inches(6.5)
                    doc.add_picture(img_path, width=max_width)
                except:
                    pass

            # ── Detect tables via filled rects ─────────────────────────────
            cell_rects = [r for r in rects if
                          r['width'] < page_w * 0.95 and
                          r['height'] < page_h * 0.3 and
                          r['width'] > 5 and r['height'] > 5]

            has_table = False
            tables = []
            y_pos = []
            x_pos = []

            if cell_rects:
                y_pos = sorted(set(
                    [round(r['top'], 1) for r in cell_rects] +
                    [round(r['bottom'], 1) for r in cell_rects]
                ))
                x_pos = sorted(set(
                    [round(r['x0'], 1) for r in cell_rects] +
                    [round(r['x1'], 1) for r in cell_rects]
                ))
                if len(y_pos) >= 2 and len(x_pos) >= 2:
                    try:
                        tables = page.extract_tables({
                            'vertical_strategy': 'explicit',
                            'horizontal_strategy': 'explicit',
                            'explicit_vertical_lines': x_pos,
                            'explicit_horizontal_lines': y_pos,
                            'snap_tolerance': 4,
                            'join_tolerance': 4,
                        }) or []
                        if tables:
                            has_table = True
                    except:
                        pass

            # Fallback: line-based detection
            if not has_table:
                try:
                    tables = page.extract_tables({
                        'vertical_strategy': 'lines',
                        'horizontal_strategy': 'lines',
                        'snap_tolerance': 3,
                    }) or []
                    if tables:
                        has_table = True
                        y_pos = []
                        x_pos = []
                except:
                    pass

            # ── CASE 1: Page has table ──────────────────────────────────────
            if has_table and tables:
                table_data = tables[0]
                rows = [r for r in table_data if any(c and str(c).strip() for c in r)]

                if rows:
                    # Merge split cells before building DOCX table
                    rows = [merge_split_cells(r) for r in rows]

                    num_cols = max(len(r) for r in rows)
                    norm_rows = [list(r) + [None] * (num_cols - len(r)) for r in rows]

                    tbl = doc.add_table(rows=len(norm_rows), cols=num_cols)
                    tbl.style = 'Table Grid'

                    row_y_spans = []
                    if len(y_pos) >= 2:
                        row_y_spans = [(y_pos[i], y_pos[i+1]) for i in range(len(y_pos)-1)]

                    col_x_spans = []
                    if len(x_pos) >= 2:
                        col_x_spans = [(x_pos[j], x_pos[j+1]) for j in range(min(num_cols, len(x_pos)-1))]

                    orig_indices = [i for i, r in enumerate(table_data)
                                    if any(c and str(c).strip() for c in r)]

                    for r_idx, row in enumerate(norm_rows):
                        orig_r = orig_indices[r_idx] if r_idx < len(orig_indices) else r_idx
                        y_top, y_bottom = row_y_spans[orig_r] if orig_r < len(row_y_spans) else (0, 0)

                        for c_idx, cell_val in enumerate(row):
                            cell = tbl.rows[r_idx].cells[c_idx]
                            text = clean_text(str(cell_val) if cell_val else '')
                            if text.lower() == 'none':
                                text = ''

                            # Background color from rects
                            if y_top != y_bottom and c_idx < len(col_x_spans):
                                cx0, cx1 = col_x_spans[c_idx]
                                hex_color = get_rect_color_at(
                                    rects, y_top, y_bottom, cx0, cx1, page_w, page_h)
                                if hex_color:
                                    try:
                                        set_cell_background(cell, hex_color)
                                    except:
                                        pass

                            if not text:
                                continue

                            rtl = is_rtl_text(text)
                            if rtl:
                                text = fix_rtl(text)

                            para = cell.paragraphs[0]
                            if rtl:
                                apply_rtl_to_paragraph(para)
                            run = para.add_run(text)
                            run.font.size = Pt(10)
                            if rtl:
                                apply_rtl_to_run(run)

                    doc.add_paragraph()

            # ── CASE 2: No table — text paragraphs ─────────────────────────
            else:
                try:
                    all_words = page.extract_words(
                        x_tolerance=3, y_tolerance=3,
                        keep_blank_chars=False,
                        use_text_flow=False,
                        extra_attrs=['fontname', 'size']
                    )
                except:
                    all_words = []

                if not all_words:
                    try:
                        import pytesseract
                        imgs = convert_from_path(input_path, dpi=150,
                                                first_page=page_index+1,
                                                last_page=page_index+1)
                        if imgs:
                            text = pytesseract.image_to_string(imgs[0], lang='eng+ara')
                            for line in text.split('\n'):
                                line = clean_text(line)
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

                lines_dict = {}
                for word in all_words:
                    y_key = round(float(word['top']) / 4) * 4
                    if y_key not in lines_dict:
                        lines_dict[y_key] = []
                    lines_dict[y_key].append(word)

                for y_key in sorted(lines_dict.keys()):
                    words = sorted(lines_dict[y_key], key=lambda w: float(w['x0']))
                    line_text = clean_text(' '.join(w['text'] for w in words))
                    if not line_text:
                        continue

                    sizes = []
                    for w in words:
                        try:
                            sizes.append(float(w.get('size', 12)))
                        except:
                            sizes.append(12.0)
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
                    run.bold = mostly_bold or (avg_size > median_size * 1.3)
                    run.font.size = Pt(min(max(8, avg_size), 72))
                    if rtl:
                        apply_rtl_to_run(run)

    doc.save(output_path)


def save_as_docx_text(input_path, output_path):
    from docx import Document
    from docx.shared import Pt, Inches

    doc = Document()

    with pdfplumber.open(input_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if i > 0:
                doc.add_page_break()

            section = doc.sections[-1]
            section.page_width = Inches(float(page.width) / 72)
            section.page_height = Inches(float(page.height) / 72)
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)

            words = page.extract_words(
                x_tolerance=3, y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
                extra_attrs=['fontname', 'size']
            )

            if not words:
                import pytesseract
                imgs = convert_from_path(input_path, dpi=150,
                                        first_page=i+1, last_page=i+1)
                if imgs:
                    text = pytesseract.image_to_string(imgs[0], lang='eng+ara')
                    for line in text.split('\n'):
                        line = clean_text(line)
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
                continue

            native_text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if native_text and native_text.strip():
                for line in native_text.split('\n'):
                    line = clean_text(line)
                    if line:
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
                    run = para.add_run(clean_text(word['text']) + ' ')
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
                save_as_docx_native(input_path, output_path, uid)

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
