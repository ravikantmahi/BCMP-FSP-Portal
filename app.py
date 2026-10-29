import os
import io
import re
import tempfile
import zipfile
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import requests
from pypdf import PdfWriter
from PIL import Image
from docxtpl import DocxTemplate

# ReportLab for fallback PDF generation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

try:
    import pythoncom
    import win32com.client
    HAS_WORD_COM = True
except ImportError:
    HAS_WORD_COM = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates') if os.path.isdir(os.path.join(BASE_DIR, 'templates')) else BASE_DIR

# DOCX Template paths (supports nomination_template.docx or template.docx)
TEMPLATE_DOCX_PATH = os.path.join(BASE_DIR, 'nomination_template.docx')
if not os.path.exists(TEMPLATE_DOCX_PATH):
    alt_path = os.path.join(BASE_DIR, 'template.docx')
    if os.path.exists(alt_path):
        TEMPLATE_DOCX_PATH = alt_path

app = Flask(__name__, template_folder=TEMPLATE_DIR)

# Google Apps Script Web App URL (Database + Drive Backend)
GAS_URL = "https://script.google.com/macros/s/AKfycbwO6x6NlvJG1l5jYp10l_SVryeXuExp5-c61GoPiIbyQpx7vsAqsq_lfTOpKbgl0TWO/exec"

# ==========================================
# Fallback Canvas (ReportLab) Configuration
# ==========================================
_FORM_IMAGE_PATH = os.path.join(BASE_DIR, 'docx_extracted', 'word', 'media', 'image1.png')

def _load_form_image_reader():
    """Load pre-rendered background image for fallback ReportLab rendering."""
    if not os.path.exists(_FORM_IMAGE_PATH):
        return None
    try:
        img = Image.open(_FORM_IMAGE_PATH).convert('RGBA')
        bg = Image.new('RGBA', img.size, (255, 255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        rgb = bg.convert('RGB')
        tmp = io.BytesIO()
        rgb.save(tmp, format='JPEG', quality=95)
        tmp.seek(0)
        return ImageReader(tmp)
    except Exception:
        return None

FORM_IMAGE_READER = _load_form_image_reader()
PAGE_W, PAGE_H = A4
FIELD_X = PAGE_W * 0.535

ROW_Y_FRACTIONS = {
    '3': 0.332, '4': 0.355, '5': 0.378, '6': 0.401, '7': 0.425,
    '8': 0.448, '9': 0.471, '10': 0.498, '11': 0.534, '12': 0.558,
    '13': 0.582, '16': 0.702, '15': 0.678,
}
CATEGORY_Y_FRACTIONS = [0.605, 0.618, 0.630, 0.643, 0.655]
CATEGORY_LABELS = [
    "Career aspirant/Student",
    "Employment Seeker/Passed out",
    "Non-IT Employees (cross-pollinated digital roles)",
    "IT Employees (IT & Non-IT sectors)",
    "PSE Employee",
]
CATEGORY_KEYWORDS = [
    ["career aspirant", "student"],
    ["employment seeker", "passed out"],
    ["non-it"],
    ["it employee", "it/pse", "it & non-it", "it &non-it"],
    ["pse employee"],
]

def frac_to_rl_y(frac):
    return PAGE_H * (1.0 - frac)

FIELD_ORDER = [
    'Timestamp', 'Title', 'Name', 'DOB', 'Gender', 'Mobile', 'Email',
    'GovID', 'Location', 'OrgName', 'Qualification', 'Status',
    'CourseYear', 'Category', 'PrevFsp', 'PrevFspDetails', 'FileUrl'
]

def get_prop(obj, *keys):
    if not isinstance(obj, dict):
        return ''
    norm = {k.lower().replace(' ', '').replace('_', '').replace('-', ''): v for k, v in obj.items()}
    for key in keys:
        nk = key.lower().replace(' ', '').replace('_', '').replace('-', '')
        if nk in norm and norm[nk] is not None:
            return str(norm[nk]).strip()
    return ''

def normalize_candidates(raw_data):
    lst = raw_data if isinstance(raw_data, list) else (raw_data.get('data') or raw_data.get('candidates') or [])
    if not lst:
        return []
    first = lst[0]
    keys = list(first.keys())
    has_headers = any(k.strip().lower() in ('name', 'applicant', 'status', 'email', 'title', 'courseyear', 'qualification') for k in keys)

    if not has_headers and len(keys) >= 8:
        reconstructed = []
        cand1 = {FIELD_ORDER[i] if i < len(FIELD_ORDER) else f'Field{i}': val for i, val in enumerate(keys)}
        reconstructed.append(cand1)
        for row in lst:
            cand = {FIELD_ORDER[i] if i < len(FIELD_ORDER) else f'Field{i}': val for i, val in enumerate(row.values())}
            reconstructed.append(cand)
        return reconstructed
    return lst

def format_dob(dob_raw):
    if not dob_raw:
        return ''
    try:
        clean = dob_raw.split('T')[0]
        dt = datetime.strptime(clean, "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    try:
        parts = dob_raw.split()
        if len(parts) >= 4:
            month_str = parts[1]
            day_str = parts[2]
            year_str = parts[3]
            dt = datetime.strptime(f"{day_str} {month_str} {year_str}", "%d %b %Y")
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return dob_raw

def build_candidate_context(candidate):
    """Build the dictionary of placeholders for the Word docx template."""
    title = get_prop(candidate, 'title')
    name = get_prop(candidate, 'name', 'applicantname', 'candidatename')
    full_name = f"{title} {name}".strip() if title else name
    dob = format_dob(get_prop(candidate, 'dob', 'dateofbirth'))
    gender = get_prop(candidate, 'gender')
    mobile = get_prop(candidate, 'mobile', 'phone')
    email = get_prop(candidate, 'email')

    loc_raw = get_prop(candidate, 'location', 'nativestate') or 'Punjab, Rupnagar (Ropar)'
    if ',' in loc_raw:
        parts = loc_raw.split(',')
        native_state = parts[0].strip()
        district = ','.join(parts[1:]).strip()
    else:
        native_state = loc_raw
        district = 'Rupnagar (Ropar)'

    gov_id = get_prop(candidate, 'govid', 'aadhar', 'gov_id')
    org_name = get_prop(candidate, 'orgname', 'organization', 'organisation', 'academicinstitute') or 'NIELIT Deemed to be University, Ropar'
    qualification = get_prop(candidate, 'qualification', 'highestqualification')
    status = get_prop(candidate, 'status')
    course_year = get_prop(candidate, 'courseyear', 'year')
    status_full = f"Pursuing ({course_year})" if (status == 'Pursuing' and course_year and course_year != 'N/A') else status

    cat_raw = get_prop(candidate, 'category').lower()
    cat_student = '✓ ' if any(kw in cat_raw for kw in ["career aspirant", "student"]) else ''
    cat_seeker = '✓ ' if any(kw in cat_raw for kw in ["employment seeker", "passed out"]) else ''
    cat_non_it = '✓ ' if any(kw in cat_raw for kw in ["non-it"]) else ''
    cat_it = '✓ ' if any(kw in cat_raw for kw in ["it employee", "it/pse", "it & non-it", "it &non-it"]) else ''
    cat_pse = '✓ ' if any(kw in cat_raw for kw in ["pse employee"]) else ''

    prev_fsp = get_prop(candidate, 'prevfsp', 'prev_fsp') or 'No'
    prev_fsp_details = get_prop(candidate, 'prevfspdetails', 'prev_fsp_details')
    if prev_fsp.lower() == 'yes' and prev_fsp_details and prev_fsp_details != 'N/A':
        fsp_lines = [l.strip() for l in prev_fsp_details.split('\n') if l.strip()]
        fsp_1 = fsp_lines[0] if fsp_lines else "N/A"
        fsp_2 = fsp_lines[1] if len(fsp_lines) > 1 else ""
    else:
        fsp_1 = "N/A"
        fsp_2 = ""

    return {
        'name': full_name,
        'applicant_name': name,
        'title': title,
        'dob': dob,
        'gender': gender,
        'mobile': mobile,
        'email': email,
        'state': native_state,
        'district': district,
        'location': f"{native_state}, {district}",
        'aadhar': gov_id,
        'govid': gov_id,
        'org_name': org_name,
        'organization': org_name,
        'qualification': qualification,
        'status': status_full,
        'status_raw': status,
        'course_year': course_year,
        'category': get_prop(candidate, 'category'),
        'cat_student': cat_student,
        'cat_seeker': cat_seeker,
        'cat_non_it': cat_non_it,
        'cat_it': cat_it,
        'cat_pse': cat_pse,
        'prev_fsp': prev_fsp,
        'prev_fsp_details': prev_fsp_details,
        'prev_fsp_1': fsp_1,
        'prev_fsp_2': fsp_2,
        'date': '2nd November 2026',
        'course_start_date': '2nd November, 2026',
    }

def convert_docx_file_to_pdf(docx_path, pdf_path, word_app=None):
    """Convert a DOCX file to PDF using Microsoft Word COM."""
    if not HAS_WORD_COM:
        raise RuntimeError("win32com is not available.")
    
    pythoncom.CoInitialize()
    should_close_word = False
    try:
        if word_app is None:
            word_app = win32com.client.DispatchEx('Word.Application')
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            should_close_word = True

        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)
        doc = word_app.Documents.Open(abs_docx, ReadOnly=True)
        # 17 = wdExportFormatPDF
        doc.SaveAs(abs_pdf, FileFormat=17)
        doc.Close(SaveChanges=0)
        return True
    finally:
        if should_close_word and word_app:
            try:
                word_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

def generate_fallback_pdf(candidate) -> bytes:
    """Fallback generator using ReportLab when Word COM is unavailable."""
    ctx = build_candidate_context(candidate)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    if FORM_IMAGE_READER:
        c.drawImage(FORM_IMAGE_READER, 0, 0, width=PAGE_W, height=PAGE_H, preserveAspectRatio=False)

    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0, 0, 0)

    simple_fields = {
        '3': ctx['name'],
        '4': ctx['dob'],
        '5': ctx['gender'],
        '6': ctx['mobile'],
        '7': ctx['email'],
        '8': ctx['state'],
        '9': ctx['district'],
        '10': ctx['aadhar'],
        '11': ctx['org_name'],
        '12': ctx['qualification'],
        '13': ctx['status'],
        '15': ctx['prev_fsp'],
        '16': f"1. {ctx['prev_fsp_1']}",
    }

    for row_key, value in simple_fields.items():
        if row_key in ROW_Y_FRACTIONS:
            y = frac_to_rl_y(ROW_Y_FRACTIONS[row_key])
            c.drawString(FIELD_X, y, str(value))

    c.drawString(FIELD_X, frac_to_rl_y(0.722), f"2. {ctx['prev_fsp_2']}")

    cat_checks = [ctx['cat_student'], ctx['cat_seeker'], ctx['cat_non_it'], ctx['cat_it'], ctx['cat_pse']]
    for i, (label, tick) in enumerate(zip(CATEGORY_LABELS, cat_checks)):
        y = frac_to_rl_y(CATEGORY_Y_FRACTIONS[i])
        mark = "✓ " if (tick and '✓' in tick) else "   "
        c.drawString(FIELD_X, y, f"{mark}{label}")

    c.save()
    buf.seek(0)
    return buf.read()

def generate_filled_pdf(candidate) -> bytes:
    """Generate filled PDF from Word template using Word COM, falling back to ReportLab if needed."""
    try:
        with tempfile.TemporaryDirectory() as td:
            docx_path = os.path.join(td, 'nomination.docx')
            pdf_path = os.path.join(td, 'nomination.pdf')
            tpl = DocxTemplate(TEMPLATE_DOCX_PATH)
            tpl.render(build_candidate_context(candidate))
            tpl.save(docx_path)

            convert_docx_file_to_pdf(docx_path, pdf_path)
            with open(pdf_path, 'rb') as f:
                return f.read()
    except Exception as e:
        print(f"[PDF Generation] Word template export error ({e}). Using ReportLab fallback.")
        return generate_fallback_pdf(candidate)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/admin')
def admin_page():
    return render_template('admin.html')

_CACHED_CANDIDATE_COUNT = 0

@app.route('/api/submit', methods=['POST'])
def api_submit():
    global _CACHED_CANDIDATE_COUNT
    try:
        payload = request.get_json(force=True)

        # 1. Determine next candidate serial number (1, 2, 3...)
        cand_idx = None
        try:
            res_c = requests.get(GAS_URL, timeout=8)
            if res_c.status_code == 200:
                cands = normalize_candidates(res_c.json())
                cand_idx = len(cands) + 1
                _CACHED_CANDIDATE_COUNT = len(cands)
        except Exception as ex:
            print(f"[api_submit] Notice: Live count query ({ex}), using cached count.")
            if _CACHED_CANDIDATE_COUNT > 0:
                _CACHED_CANDIDATE_COUNT += 1
                cand_idx = _CACHED_CANDIDATE_COUNT

        # 2. Rename uploaded ID document to: {index}_{candidate_name}.{ext}
        orig_filename = payload.get('fileName', '')
        cand_name = payload.get('name', '').strip()

        if orig_filename and cand_name:
            # Enforce .pdf extension for Drive upload
            safe_name = re.sub(r'[\\/*?:"<>|]', '', cand_name).strip()
            safe_name = " ".join(safe_name.split())
            if not safe_name:
                safe_name = "Candidate"

            new_filename = f"{cand_idx}_{safe_name}.pdf" if cand_idx is not None else f"{safe_name}.pdf"

            payload['fileName'] = new_filename
            payload['fileMimeType'] = 'application/pdf'
            print(f"[api_submit] Renamed uploaded file for Drive: {new_filename}")

        res = requests.post(GAS_URL, json=payload, headers={'Content-Type': 'text/plain;charset=utf-8'}, timeout=60)
        
        # If successfully submitted, increment cached count
        if cand_idx is not None:
            _CACHED_CANDIDATE_COUNT = cand_idx

        return jsonify(res.json())
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/candidates', methods=['GET'])
def api_candidates():
    global _CACHED_CANDIDATE_COUNT
    try:
        res = requests.get(GAS_URL, timeout=30)
        raw_data = res.json()
        if isinstance(raw_data, dict) and raw_data.get('error'):
            return jsonify({'error': raw_data['error']}), 500
        cands = normalize_candidates(raw_data)
        _CACHED_CANDIDATE_COUNT = len(cands)
        return jsonify(cands)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download-single/<int:idx>', methods=['GET'])
def download_single(idx):
    try:
        res = requests.get(GAS_URL, timeout=30)
        candidates = normalize_candidates(res.json())
        if idx < 0 or idx >= len(candidates):
            return "Candidate not found", 404

        c = candidates[idx]
        pdf_bytes = io.BytesIO(generate_filled_pdf(c))

        name_clean = "".join(ch if ch.isalnum() else "_" for ch in (get_prop(c, 'name') or 'Candidate'))
        year_clean = (get_prop(c, 'courseyear') or 'N_A').replace(' ', '')
        return send_file(pdf_bytes, mimetype='application/pdf', as_attachment=True,
                         download_name=f"{name_clean}_{year_clean}_AVR_Bootcamp.pdf")
    except Exception as e:
        return f"Error generating PDF: {str(e)}", 500

@app.route('/api/download-single-docx/<int:idx>', methods=['GET'])
def download_single_docx(idx):
    try:
        res = requests.get(GAS_URL, timeout=30)
        candidates = normalize_candidates(res.json())
        if idx < 0 or idx >= len(candidates):
            return "Candidate not found", 404

        c = candidates[idx]
        tpl = DocxTemplate(TEMPLATE_DOCX_PATH)
        tpl.render(build_candidate_context(c))

        docx_buffer = io.BytesIO()
        tpl.save(docx_buffer)
        docx_buffer.seek(0)

        name_clean = "".join(ch if ch.isalnum() else "_" for ch in (get_prop(c, 'name') or 'Candidate'))
        year_clean = (get_prop(c, 'courseyear') or 'N_A').replace(' ', '')
        return send_file(
            docx_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f"{name_clean}_{year_clean}_AVR_Bootcamp.docx"
        )
    except Exception as e:
        return f"Error generating DOCX: {str(e)}", 500

@app.route('/api/download-merged', methods=['GET'])
def download_merged():
    year = request.args.get('year', 'All')
    try:
        res = requests.get(GAS_URL, timeout=30)
        candidates = normalize_candidates(res.json())
        if year and year != 'All':
            candidates = [c for c in candidates if get_prop(c, 'courseyear', 'year') == year]

        if not candidates:
            return f"No candidates found for {year}", 404

        merger = PdfWriter()

        word_app = None
        use_fallback = False
        try:
            if HAS_WORD_COM:
                pythoncom.CoInitialize()
                word_app = win32com.client.DispatchEx('Word.Application')
                word_app.Visible = False
                word_app.DisplayAlerts = 0
            else:
                use_fallback = True
        except Exception as e:
            print(f"[Merged PDF] Word COM init failed ({e}), using ReportLab fallback.")
            use_fallback = True

        with tempfile.TemporaryDirectory() as td:
            for i, c in enumerate(candidates):
                if not use_fallback and word_app:
                    try:
                        docx_path = os.path.join(td, f"cand_{i}.docx")
                        pdf_path = os.path.join(td, f"cand_{i}.pdf")
                        tpl = DocxTemplate(TEMPLATE_DOCX_PATH)
                        tpl.render(build_candidate_context(c))
                        tpl.save(docx_path)
                        convert_docx_file_to_pdf(docx_path, pdf_path, word_app=word_app)
                        merger.append(pdf_path)
                        continue
                    except Exception as err:
                        print(f"[Merged PDF] Candidate {i} Word error: {err}")
                # Fallback to reportlab
                merger.append(io.BytesIO(generate_fallback_pdf(c)))

        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

        merged_buffer = io.BytesIO()
        merger.write(merged_buffer)
        merger.close()
        merged_buffer.seek(0)

        safe_year = 'All_Candidates' if year == 'All' else year.replace(' ', '')
        return send_file(merged_buffer, mimetype='application/pdf', as_attachment=True,
                         download_name=f"AVR_Bootcamp_Merged_{safe_year}.pdf")
    except Exception as e:
        return f"Error merging PDFs: {str(e)}", 500

@app.route('/api/download-zip', methods=['GET'])
def download_zip():
    year = request.args.get('year', 'All')
    try:
        res = requests.get(GAS_URL, timeout=30)
        candidates = normalize_candidates(res.json())
        if year and year != 'All':
            candidates = [c for c in candidates if get_prop(c, 'courseyear', 'year') == year]

        if not candidates:
            return f"No candidates found for {year}", 404

        word_app = None
        use_fallback = False
        try:
            if HAS_WORD_COM:
                pythoncom.CoInitialize()
                word_app = win32com.client.DispatchEx('Word.Application')
                word_app.Visible = False
                word_app.DisplayAlerts = 0
            else:
                use_fallback = True
        except Exception as e:
            print(f"[ZIP PDF] Word COM init failed ({e}), using ReportLab fallback.")
            use_fallback = True

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            with tempfile.TemporaryDirectory() as td:
                for i, c in enumerate(candidates):
                    name_clean = "".join(ch if ch.isalnum() else "_" for ch in (get_prop(c, 'name') or f'Candidate_{i+1}'))
                    year_clean = (get_prop(c, 'courseyear') or year).replace(' ', '')
                    filename = f"{name_clean}_{year_clean}_AVR_Bootcamp.pdf"

                    pdf_bytes = None
                    if not use_fallback and word_app:
                        try:
                            docx_path = os.path.join(td, f"cand_{i}.docx")
                            pdf_path = os.path.join(td, f"cand_{i}.pdf")
                            tpl = DocxTemplate(TEMPLATE_DOCX_PATH)
                            tpl.render(build_candidate_context(c))
                            tpl.save(docx_path)
                            convert_docx_file_to_pdf(docx_path, pdf_path, word_app=word_app)
                            with open(pdf_path, 'rb') as pf:
                                pdf_bytes = pf.read()
                        except Exception as err:
                            print(f"[ZIP PDF] Candidate {i} Word error: {err}")

                    if not pdf_bytes:
                        pdf_bytes = generate_fallback_pdf(c)

                    zf.writestr(filename, pdf_bytes)

        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

        zip_buffer.seek(0)
        safe_year = 'All' if year == 'All' else year.replace(' ', '')
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True,
                         download_name=f"AVR_Nominations_{safe_year}.zip")
    except Exception as e:
        return f"Error creating ZIP: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
