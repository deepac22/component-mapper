import streamlit as st
import pandas as pd
import re
import io
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Page config
st.set_page_config(page_title="Flex PCBA Mapper", page_icon="🔧", layout="wide")

# Header with logo and title
st.markdown(
    """
    <div style="display:flex; align-items:center; gap:15px; margin-bottom:20px;">
        <svg width="140" viewBox="0 0 280 80">
            <text x="10" y="56" font-family="Arial" font-size="52" font-weight="700" fill="#0077C8">flex</text>
            <text x="14" y="72" font-family="Arial" font-size="11" font-weight="500" letter-spacing="3.5" fill="#8899aa">PCBA COMPONENT MAPPER</text>
        </svg>
        <div style="font-size:1.2rem; font-weight:600; color:#2c3e50;">
            <span style="color:#0077C8;">PCBA</span> Component Mapping Tool
        </div>
    </div>
    """,
    unsafe_allow_html=True
)
st.markdown("Upload your **PCBA Assembly Drawing** and **BOM Excel** to automatically map components and download a report.")

# --- Helper functions ---

# Set Tesseract path (Streamlit Cloud uses Linux)
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

COMP_PREFIXES = ['R','C','L','D','Q','U','J','TP','X','Y','FB','T','F','P','K','S','W']

def extract_text_from_image(image_bytes):
    image = Image.open(io.BytesIO(image_bytes)).convert('RGB')
    return pytesseract.image_to_string(image)

def extract_text_from_pdf(pdf_bytes):
    pages = convert_from_bytes(pdf_bytes, dpi=300)
    text = ""
    for page in pages:
        text += pytesseract.image_to_string(page) + "\n"
    return text

def extract_component_codes(text):
    codes = set()
    prefix_pattern = '|'.join(COMP_PREFIXES)
    pattern = re.compile(r'\b(' + prefix_pattern + r')(\d+[A-Za-z]?)\b', re.IGNORECASE)
    for match in pattern.finditer(text):
        codes.add(match.group().upper())
    return sorted(codes)

def parse_bom(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    comp_col = part_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if comp_col is None and re.search(r'component\s*code|ref\s*des|designator|item\s*code|code', col_lower):
            comp_col = col
        if part_col is None and re.search(r'part\s*no|part\s*number|pn|part\s*#|manufacturer\s*part', col_lower):
            part_col = col
    if comp_col is None:
        comp_col = df.columns[0] if len(df.columns) > 0 else None
    if part_col is None:
        part_col = df.columns[1] if len(df.columns) > 1 else None
    if comp_col is None or part_col is None:
        return {}
    mapping = {}
    for _, row in df.iterrows():
        code = str(row[comp_col]).strip()
        part = str(row[part_col]).strip()
        if code and part and code.lower() != 'nan' and part.lower() != 'nan':
            mapping[code.upper()] = part
    return mapping

def generate_excel(results):
    wb = Workbook()
    ws = wb.active
    ws.title = "PCBA Mapping"
    header_fill = PatternFill(start_color="0077C8", end_color="0077C8", fill_type="solid")
    headers = ["PartNo", "Component code in drawing", "Mapping code"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for row_idx, item in enumerate(results, 2):
        ws.cell(row=row_idx, column=1, value=item['part_no'])
        ws.cell(row=row_idx, column=2, value=item['component_code_in_drawing'])
        ws.cell(row=row_idx, column=3, value=item['mapping_code'])
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Main UI ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📐 PCBA Assembly Drawing")
    drawing_file = st.file_uploader(
        "Upload the assembly drawing",
        type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp"],
        key="drawing"
    )

with col2:
    st.subheader("📊 BOM Excel File")
    bom_file = st.file_uploader(
        "Upload the Bill of Materials",
        type=["xlsx", "xls"],
        key="bom"
    )

if drawing_file and bom_file:
    if st.button("⚡ Process & Generate Report", type="primary", use_container_width=True):
        with st.spinner("Processing drawing and BOM... Please wait."):
            # --- Extract text from drawing ---
            drawing_bytes = drawing_file.read()
            fname = drawing_file.name.lower()
            try:
                if fname.endswith('.pdf'):
                    full_text = extract_text_from_pdf(drawing_bytes)
                else:
                    full_text = extract_text_from_image(drawing_bytes)
            except Exception as e:
                st.error(f"Failed to process drawing: {e}")
                st.stop()

            # --- Extract component codes ---
            codes = extract_component_codes(full_text)

            # --- Parse BOM ---
            bom_bytes = bom_file.read()
            try:
                bom_map = parse_bom(bom_bytes)
            except Exception as e:
                st.error(f"Failed to parse BOM Excel: {e}")
                st.stop()

            # --- Map ---
            results = []
            for code in codes:
                part = bom_map.get(code, "")
                results.append({
                    'part_no': part,
                    'component_code_in_drawing': code,
                    'mapping_code': code if part else "",
                    'status': 'MATCHED' if part else 'NOT_IN_BOM'
                })

            st.success("✅ Processing complete!")

            # --- Display stats ---
            matched = sum(1 for r in results if r['status'] == 'MATCHED')
            not_found = len(results) - matched
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("🔍 Drawing Components", len(codes))
            col_m2.metric("✅ Matched with BOM", matched)
            col_m3.metric("❌ Not Found in BOM", not_found)

            # --- Show table ---
            st.subheader("📋 Mapping Table")
            df_results = pd.DataFrame(results)
            # Reorder columns
            df_display = df_results[['part_no', 'component_code_in_drawing', 'mapping_code', 'status']].copy()
            df_display.columns = ["PartNo", "Component Code in Drawing", "Mapping Code", "Status"]
            st.dataframe(df_display, use_container_width=True, hide_index=True)

            # --- Generate Excel and download button ---
            excel_data = generate_excel(results)
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name="pcba_mapping_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # Show extracted raw text (optional, hidden by default)
            with st.expander("🔎 View extracted text from drawing (for debugging)"):
                st.text_area("OCR Output", full_text, height=200)

else:
    st.info("👆 Please upload both files to start mapping.")
