import os
import re
import io
import tempfile
from flask import Flask, request, render_template, send_file, jsonify
import pandas as pd
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max

# Common component reference prefixes
COMP_PREFIXES = {'R', 'C', 'L', 'D', 'Q', 'U', 'J', 'TP', 'X', 'Y', 'FB', 'T', 'F', 'P', 'K', 'S', 'W'}

def extract_text_from_image(image_bytes):
    """Extract text from image bytes using OCR."""
    image = Image.open(io.BytesIO(image_bytes))
    # Convert to RGB if necessary
    if image.mode != 'RGB':
        image = image.convert('RGB')
    text = pytesseract.image_to_string(image)
    return text

def extract_text_from_pdf(pdf_bytes):
    """Extract text from PDF by converting each page to image."""
    pages = convert_from_bytes(pdf_bytes, dpi=300)
    full_text = ""
    for page in pages:
        text = pytesseract.image_to_string(page)
        full_text += text + "\n"
    return full_text

def extract_component_codes(text):
    """
    Extract component reference designators from OCR text.
    e.g., R1, C10, U3, J2, etc.
    """
    codes = set()
    # Build regex: prefix (case insensitive) followed by digits, optionally with suffix like A, B...
    prefix_pattern = '|'.join(COMP_PREFIXES)
    pattern = re.compile(r'\b(' + prefix_pattern + r')(\d+[A-Za-z]?)\b', re.IGNORECASE)
    for line in text.splitlines():
        matches = pattern.findall(line)
        for prefix, num in matches:
            codes.add(prefix.upper() + num.upper())
    return sorted(codes)

def parse_bom(file_bytes, filename):
    """
    Parse BOM Excel file and return a dict mapping component code -> part number.
    Automatically detects columns.
    """
    df = pd.read_excel(io.BytesIO(file_bytes))
    # Try to find column names for component code and part number
    comp_col = None
    part_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if comp_col is None and re.search(r'component\s*code|ref\s*des|designator|item\s*code|code', col_lower):
            comp_col = col
        if part_col is None and re.search(r'part\s*no|part\s*number|pn|part\s*#|manufacturer\s*part', col_lower):
            part_col = col
    # Fallback: assume first two columns if not detected
    if comp_col is None:
        comp_col = df.columns[0] if len(df.columns) > 0 else None
    if part_col is None:
        part_col = df.columns[1] if len(df.columns) > 1 else None
    if comp_col is None or part_col is None:
        return {}
    # Create mapping dictionary (keep first match if duplicates)
    mapping = {}
    for _, row in df.iterrows():
        code = str(row[comp_col]).strip()
        part = str(row[part_col]).strip()
        if code and part and code.lower() != 'nan' and part.lower() != 'nan':
            # Normalize code for matching
            mapping[code.upper()] = part
    return mapping

def generate_excel_report(results):
    """Generate an Excel file with the report."""
    wb = Workbook()
    ws = wb.active
    ws.title = "PCBA Mapping Report"
    # Header style
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="0077C8", end_color="0077C8", fill_type="solid")
    headers = ["PartNo", "Component code in drawing", "Mapping code"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    # Data rows
    for row_idx, item in enumerate(results, 2):
        ws.cell(row=row_idx, column=1, value=item.get('part_no', ''))
        ws.cell(row=row_idx, column=2, value=item.get('component_code_in_drawing', ''))
        ws.cell(row=row_idx, column=3, value=item.get('mapping_code', ''))
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 4
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    if 'drawing' not in request.files or 'bom' not in request.files:
        return jsonify({'error': 'Missing drawing or BOM file'}), 400

    drawing_file = request.files['drawing']
    bom_file = request.files['bom']

    # --- Extract text from drawing ---
    drawing_bytes = drawing_file.read()
    filename = drawing_file.filename.lower()
    try:
        if filename.endswith('.pdf'):
            full_text = extract_text_from_pdf(drawing_bytes)
        else:
            full_text = extract_text_from_image(drawing_bytes)
    except Exception as e:
        return jsonify({'error': f'Failed to process drawing: {str(e)}'}), 500

    # --- Extract component codes ---
    drawing_codes = extract_component_codes(full_text)

    # --- Parse BOM ---
    bom_bytes = bom_file.read()
    try:
        bom_mapping = parse_bom(bom_bytes, bom_file.filename)
    except Exception as e:
        return jsonify({'error': f'Failed to parse BOM Excel: {str(e)}'}), 500

    # --- Map components ---
    results = []
    for code in drawing_codes:
        part_no = bom_mapping.get(code, "")
        if part_no:
            mapping_code = code  # Mapping code is the same as drawing code (matched)
            status = 'MATCHED'
        else:
            # Check if a similar code exists in BOM (optional fuzzy matching, skipped for simplicity)
            mapping_code = ""
            status = 'NOT_IN_BOM'
        results.append({
            'part_no': part_no,
            'component_code_in_drawing': code,
            'mapping_code': mapping_code,
            'status': status
        })

    # --- Generate Excel report ---
    report_io = generate_excel_report(results)
    # Save to a temporary file to send as download, or store in memory and provide a download link.
    # For simplicity, we'll save to a temp file and send it directly in a download endpoint.
    # Here we'll return a JSON with the download URL (simplified by storing in a global dict).
    # Better: save to a temp dir and generate unique ID.
    temp_dir = tempfile.gettempdir()
    report_path = os.path.join(temp_dir, 'pcba_mapping_report.xlsx')
    with open(report_path, 'wb') as f:
        f.write(report_io.getbuffer())
    # We'll provide a download URL (using the same route)
    return jsonify({
        'success': True,
        'results': results,
        'total_drawing_components': len(drawing_codes),
        'total_bom_entries': len(bom_mapping),
        'download_url': '/download_report'
    })

@app.route('/download_report')
def download_report():
    temp_dir = tempfile.gettempdir()
    report_path = os.path.join(temp_dir, 'pcba_mapping_report.xlsx')
    if os.path.exists(report_path):
        return send_file(report_path, as_attachment=True, download_name='pcba_mapping_report.xlsx')
    else:
        return "Report not found. Please process first.", 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
