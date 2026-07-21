"""
Flex PCBA CompMap – Component Mapping Web App
Supports: PDF (text + OCR fallback), DXF (blocks included), PNG/JPG/TIFF/BMP
Matches components using either a direct BOM Excel (Component Code → Part No)
or a category reference Excel (Code → Description).
"""

import io
import re
import numpy as np
import cv2
import pytesseract
from pytesseract import Output
import fitz   # PyMuPDF
import ezdxf
import openpyxl
from openpyxl.styles import Font
import streamlit as st
import pandas as pd
from PIL import Image

# -------------------------------------------------------------------
# Styling & page setup
# -------------------------------------------------------------------
st.set_page_config(page_title="CompMap | Flex", page_icon="◆", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; }
    .hero {
        padding: 2.2rem 2rem; border-radius: 16px;
        background: linear-gradient(135deg, #0091D5 0%, #005A8C 100%);
        margin-bottom: 1.8rem; box-shadow: 0 8px 32px rgba(0, 145, 213, 0.25);
        display: flex; justify-content: space-between; align-items: center;
    }
    .hero-left h1 { color: white; font-size: 2rem; font-weight: 700; margin: 0; letter-spacing: -0.02em; }
    .hero-left p { color: rgba(255,255,255,0.9); margin-top: 0.4rem; font-size: 0.95rem; }
    .badge {
        display: inline-block; background: rgba(255,255,255,0.2); color: white;
        padding: 3px 12px; border-radius: 20px; font-size: 0.7rem;
        font-family: 'JetBrains Mono', monospace; letter-spacing: 0.05em; margin-top: 0.6rem;
    }
    .flex-wordmark {
        font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 1.4rem;
        color: #0091D5; background: white; letter-spacing: 0.02em;
        padding: 8px 20px; border-radius: 8px;
    }
    div[data-testid="stMetric"] { background: #F0F8FC; border: 1px solid #B3E0F5; border-radius: 12px; padding: 1rem 1.2rem; }
    div[data-testid="stMetricValue"] { font-family: 'JetBrains Mono', monospace; color: #0091D5; }
    .stDataFrame { border-radius: 12px; overflow: hidden; }
    div[data-testid="stFileUploader"] { border: 1px dashed #0091D5; border-radius: 12px; padding: 0.5rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <div class="hero-left">
        <h1>◆ CompMap</h1>
        <p>Upload a PCBA assembly drawing (PDF, DXF, or image) and a reference file — get a mapping report.</p>
        <span class="badge">v2.2 · universal extraction · BOM or category matching</span>
    </div>
    <div class="flex-wordmark">flex</div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------
# Patterns and helpers
# -------------------------------------------------------------------
# Generic component label: 1-4 letters, followed by 1-4 digits, optional suffix letter
LABEL_PATTERN = re.compile(r'^[A-Z]{1,4}\d{1,4}[A-Z]?$')

# Legacy CO prefix pattern (still supported)
CO_PATTERN = re.compile(r'\bCO([A-Z]{1,4}\d+[A-Z]?)\b')

# Standard designator: R1, C10, U3, SW2, etc.
STANDARD_CODE_PATTERN = re.compile(r'\b([A-Z]{1,6})\s*(\d{1,4}[A-Za-z]?)\b')

# Tesseract path for Streamlit Cloud
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

# -------------------------------------------------------------------
# OCR with adaptive preprocessing
# -------------------------------------------------------------------
def ocr_grayscale_image(gray_image, grid=6):
    """Extract component labels using grid-based OCR with contrast enhancement."""
    found = set()
    h, w = gray_image.shape

    # Preprocessing: CLAHE (contrast enhancement) and Otsu threshold
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    gray_clahe = clahe.apply(gray_image)
    _, thresh_otsu = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    for img in [gray_clahe, thresh_otsu]:
        for r in range(grid):
            for c in range(grid):
                y1 = max(0, r * h // grid - 20)
                y2 = min(h, (r + 1) * h // grid + 20)
                x1 = max(0, c * w // grid - 20)
                x2 = min(w, (c + 1) * w // grid + 20)
                tile = img[y1:y2, x1:x2]

                # Try different page segmentation modes:
                # 6 = uniform block of text, 11 = sparse text, 3 = fully automatic (rotated)
                for psm in [6, 11, 3]:
                    data = pytesseract.image_to_data(tile, output_type=Output.DICT,
                                                     config=f'--psm {psm}')
                    for i in range(len(data['text'])):
                        t = data['text'][i].strip().upper()
                        conf = int(data['conf'][i]) if data['conf'][i] != '-1' else -1
                        if conf > 30 and LABEL_PATTERN.match(t):
                            found.add(t)
    return found

# -------------------------------------------------------------------
# PDF extraction (text first, OCR fallback)
# -------------------------------------------------------------------
def extract_from_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    all_codes = set()
    total_text_length = 0

    for page in doc:
        text = page.get_text()
        total_text_length += len(text)

        # Legacy CO pattern
        for m in CO_PATTERN.finditer(text):
            all_codes.add(m.group(1).upper())

        # Standard designators (R1, C2, ...)
        for match in STANDARD_CODE_PATTERN.finditer(text):
            prefix, num = match.groups()
            prefix = prefix.upper()
            # Accept only if prefix is all letters (avoid matching e.g. "12V")
            if prefix.isalpha() and len(prefix) <= 4:
                candidate = prefix + num.upper()
                if LABEL_PATTERN.match(candidate):
                    all_codes.add(candidate)

    if len(all_codes) > 0:
        return sorted(all_codes), "text", None

    # If very little text, assume scanned PDF → OCR fallback
    if total_text_length < 50:
        ocr_found = set()
        for page in doc:
            pix = page.get_pixmap(dpi=400)
            img_array = np.frombuffer(pix.tobytes("png"), dtype=np.uint8)
            img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            ocr_found |= ocr_grayscale_image(gray)
        if len(ocr_found) == 0:
            return [], None, "No embedded text, and OCR found no component labels either."
        return sorted(ocr_found), "ocr", None

    return [], None, "PDF has text, but no component labels could be identified."

# -------------------------------------------------------------------
# DXF extraction (block‑aware)
# -------------------------------------------------------------------
def extract_from_dxf(file_bytes):
    try:
        text_stream = io.StringIO(file_bytes.decode('utf-8', errors='ignore'))
        doc = ezdxf.read(text_stream)
    except Exception as e:
        return [], f"Couldn't read this DXF file. ({e})"

    msp = doc.modelspace()
    found = set()

    def get_text(entity):
        """Recursively extract text from entity, including blocks."""
        if entity.dxftype() in ('TEXT', 'ATTRIB'):
            return entity.dxf.text
        elif entity.dxftype() == 'MTEXT':
            return entity.text
        elif entity.dxftype() == 'INSERT':
            block = doc.blocks.get(entity.dxf.name)
            if block:
                for sub_entity in block:
                    txt = get_text(sub_entity)
                    if txt:
                        return txt
        return None

    for entity in msp:
        txt = get_text(entity)
        if txt and LABEL_PATTERN.match(txt.strip().upper()):
            found.add(txt.strip().upper())

    if len(found) == 0:
        return [], "No component labels found in DXF (including block contents)."
    return sorted(found), None

# -------------------------------------------------------------------
# Image extraction (OCR only)
# -------------------------------------------------------------------
def extract_from_image(file_bytes):
    try:
        pil_image = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        img_array = np.array(pil_image)
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    except Exception as e:
        return [], f"Couldn't open this image file. ({e})"

    found = ocr_grayscale_image(gray)
    if len(found) == 0:
        return [], "OCR couldn't find any component labels in this image."
    return sorted(found), None

# -------------------------------------------------------------------
# Main dispatcher
# -------------------------------------------------------------------
def extract_component_codes(uploaded_file):
    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.read()

    if filename.endswith('.pdf'):
        codes, method, error = extract_from_pdf(file_bytes)
        return codes, method, error

    elif filename.endswith('.dxf'):
        codes, error = extract_from_dxf(file_bytes)
        return codes, "dxf-text", error

    elif filename.endswith(('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp')):
        codes, error = extract_from_image(file_bytes)
        return codes, "ocr", error

    elif filename.endswith('.dwg'):
        return [], None, "**.dwg files are not supported.** Please re-export as PDF or DXF."

    else:
        return [], None, f"Unsupported file type: {filename.split('.')[-1] if '.' in filename else 'unknown'}"

# -------------------------------------------------------------------
# Excel loading – BOM vs. Category Reference
# -------------------------------------------------------------------
def load_bom_table(xlsx_bytes):
    """Detect columns: Component Code + Part No → return dict for direct mapping.
    Returns None if columns not found."""
    try:
        df = pd.read_excel(io.BytesIO(xlsx_bytes))
    except Exception:
        return None

    comp_col = part_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if comp_col is None and re.search(r'component\s*code|ref\s*des|designator|item\s*code|code|reference', col_lower):
            comp_col = col
        if part_col is None and re.search(r'part\s*no|part\s*number|pn|part\s*#|manufacturer\s*part', col_lower):
            part_col = col
    if comp_col is None or part_col is None:
        return None

    mapping = {}
    for _, row in df.iterrows():
        code = str(row[comp_col]).strip().upper()
        part = str(row[part_col]).strip()
        if code and part and code.lower() != 'nan' and part.lower() != 'nan':
            mapping[code] = part
    return mapping

def load_category_reference(xlsx_bytes):
    """Detect columns: Code + Description → list of (code, description).
    Returns (entries, error_message)."""
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
        sheet = workbook.active
    except Exception as e:
        return [], f"Couldn't open this Excel file. ({e})"

    rows = list(sheet.iter_rows(min_row=1, values_only=True))
    if not rows:
        return [], "This reference file appears to be empty."

    header_row = [str(h).strip().lower() if h else "" for h in rows[0]]
    code_col = desc_col = None
    for i, h in enumerate(header_row):
        if "code" in h and code_col is None:
            code_col = i
        if "desc" in h:
            desc_col = i
    if code_col is None:
        code_col = 0
    if desc_col is None:
        desc_col = 2 if len(header_row) > 2 else len(header_row) - 1

    entries = []
    for row in rows[1:]:
        if len(row) <= max(code_col, desc_col):
            continue
        code, desc = row[code_col], row[desc_col]
        if code and desc:
            entries.append((code, str(desc)))

    if not entries:
        return [], "Couldn't find usable code/description rows. Expected headers containing 'Code' and 'Description'."
    return entries, None

# -------------------------------------------------------------------
# Category mapping helpers (when not using direct BOM)
# -------------------------------------------------------------------
def get_prefix(component_code):
    match = re.match(r'^([A-Z]+)\d', component_code)
    return match.group(1) if match else component_code

PREFIX_TO_KEYWORDS = {
    'R': ['RESISTOR'], 'C': ['CAPACITOR'], 'L': ['INDUCTOR'],
    'D': ['DIODE', 'RECTIFIER'], 'Q': ['TRANSISTOR'], 'J': ['CONNECTOR'],
    'SW': ['SWITCH'], 'F': ['FUSE'], 'NT': ['ANTENNA'], 'ANT': ['ANTENNA'],
    'X': ['CRYSTAL', 'OSCILLATOR'], 'Y': ['CRYSTAL', 'OSCILLATOR'],
}
NEEDS_REVIEW_PREFIXES = {'U', 'MP', 'TP', 'PTH', 'FD', 'FILT', 'DMC', 'Z', 'P', 'A'}

def find_match_in_reference(keywords, reference_entries):
    """Pick the most generic matching description (fewest extra words)."""
    best_match = None
    best_extra_words = None
    for keyword in keywords:
        for code, description in reference_entries:
            words = re.findall(r'[A-Za-z]+', description.upper())
            matching_words = [w for w in words if keyword in w]
            if not matching_words:
                continue
            extra = len(words) - len(matching_words)
            if best_extra_words is None or extra < best_extra_words:
                best_extra_words = extra
                best_match = (code, description)
        if best_match:
            break
    return best_match

def map_component_category(component_code, reference_entries):
    prefix = get_prefix(component_code)
    if prefix in PREFIX_TO_KEYWORDS:
        match = find_match_in_reference(PREFIX_TO_KEYWORDS[prefix], reference_entries)
        if match:
            return match[0], match[1], "OK"
        return "NOT IN FILE", f"No '{PREFIX_TO_KEYWORDS[prefix][0]}' category found", "REVIEW"
    elif prefix in NEEDS_REVIEW_PREFIXES:
        return "NEEDS REVIEW", f"'{prefix}' needs manual classification", "REVIEW"
    else:
        return "UNKNOWN PREFIX", f"'{prefix}' is not recognized", "REVIEW"

# -------------------------------------------------------------------
# Excel report builder
# -------------------------------------------------------------------
def build_excel_bytes(rows, mode='bom'):
    """rows is a list of dicts with keys: Partno, Component code in the drawing,
    Mapping Code, Description, Status."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    headers = ["Partno", "Component code in the drawing", "Mapping Code", "Description", "Status"]
    for col_num, header in enumerate(headers, 1):
        sheet.cell(row=1, column=col_num, value=header).font = Font(bold=True)
    for row_idx, row in enumerate(rows, 2):
        sheet.cell(row=row_idx, column=1, value=row['Partno'])
        sheet.cell(row=row_idx, column=2, value=row['Component code in the drawing'])
        sheet.cell(row=row_idx, column=3, value=row['Mapping Code'])
        sheet.cell(row=row_idx, column=4, value=row.get('Description', ''))
        sheet.cell(row=row_idx, column=5, value=row['Status'])
    for col, width in {"A": 8, "B": 30, "C": 16, "D": 45, "E": 12}.items():
        sheet.column_dimensions[col].width = width
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer

# -------------------------------------------------------------------
# Main UI
# -------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    st.markdown("**📄 Assembly Drawing** — PDF, DXF, PNG, JPG, TIFF, or BMP")
    drawing_file = st.file_uploader(
        "Drawing", type=["pdf", "dxf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "dwg"],
        label_visibility="collapsed"
    )
with col2:
    st.markdown("**📋 Reference File** (.xlsx)")
    ref_file = st.file_uploader("Excel", type=["xlsx"], label_visibility="collapsed")

if drawing_file and ref_file:
    with st.spinner("Extracting and mapping components..."):
        codes, method, drawing_error = extract_component_codes(drawing_file)

    if drawing_error:
        st.error(f"**Drawing issue:** {drawing_error}")

    if not drawing_error:
        ref_bytes = ref_file.read()

        # --- Auto‑detect BOM mode ---
        bom_map = load_bom_table(ref_bytes)
        if bom_map is not None:
            st.info("📦 **BOM mode** – matching component codes to part numbers directly.")
            rows = []
            for i, code in enumerate(codes, start=1):
                part = bom_map.get(code, "")
                rows.append({
                    "Partno": i,
                    "Component code in the drawing": code,
                    "Mapping Code": part if part else "",
                    "Description": "",
                    "Status": "MATCHED" if part else "NOT_IN_BOM"
                })
            matched = sum(1 for r in rows if r['Status'] == 'MATCHED')
            not_found = len(rows) - matched

        else:
            # --- Category reference mode ---
            reference_entries, ref_error = load_category_reference(ref_bytes)
            if ref_error:
                st.error(f"**Reference file issue:** {ref_error}")
                st.stop()
            st.info("📚 **Category mode** – mapping by component type (resistor, capacitor, etc.).")
            rows = []
            for i, code in enumerate(codes, start=1):
                mcode, desc, status = map_component_category(code, reference_entries)
                rows.append({
                    "Partno": i,
                    "Component code in the drawing": code,
                    "Mapping Code": mcode,
                    "Description": desc,
                    "Status": status
                })
            matched = sum(1 for r in rows if r['Status'] == 'OK')
            not_found = len(rows) - matched

        # Show extraction method warning
        if method == "ocr":
            st.warning(
                "⚠️ **This file has no embedded text** — results come from OCR, "
                "which may be less accurate. Use PDF or DXF with real text when possible."
            )
        elif method == "dxf-text":
            st.success("✓ Read directly from DXF text entities.")

        # Metrics
        total = len(rows)
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total components", total)
        col_m2.metric("Matched / OK", matched)
        col_m3.metric("Unmatched / Review", not_found)

        # Table
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, height=400)

        # Download
        excel_bytes = build_excel_bytes(rows)
        st.download_button(
            "⬇ Download Excel Report",
            data=excel_bytes,
            file_name="Component_Mapping_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Upload both files above to generate your report.")
    
