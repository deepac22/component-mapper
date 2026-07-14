

import re
import io
import fitz
import openpyxl
from openpyxl.styles import Font
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

def extract_component_codes(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pattern = re.compile(r'\bCO([A-Z]{1,4}\d+[A-Z]?)\b')
    all_codes = set()
    for page in doc:
        all_codes.update(pattern.findall(page.get_text()))
    return sorted(all_codes)


def load_reference_table(xlsx_bytes):
    workbook = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    sheet = workbook.active
    lookup = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        code, group, description = row
        if code:
            lookup[code] = description
    return lookup


def get_prefix(component_code):
    match = re.match(r'^([A-Z]+)\d', component_code)
    return match.group(1) if match else component_code


PREFIX_TO_COMMODITY_CODE = {
    'R': 'PR000', 'C': 'PC000', 'L': 'PI000',
    'D': 'LDD51', 'Q': 'LDT30', 'J': '11000',
    'SW': 'SW000', 'F': 'PP100', 'NT': '43000',
    'ANT': '43000', 'X': 'PF400', 'Y': 'PF400',
}
NEEDS_REVIEW_PREFIXES = {'U', 'MP', 'TP', 'PTH', 'FD', 'FILT', 'DMC', 'Z'}
CATEGORY_NAMES = {
    'PR000': 'Resistors', 'PC000': 'Capacitors', 'PI000': 'Inductors',
    'LDD51': 'Diodes', 'LDT30': 'Transistors', '11000': 'Connectors',
    'SW000': 'Switches', 'PP100': 'Fuses', '43000': 'Antennas',
    'PF400': 'Crystals/Osc',
}


def map_component(component_code, reference_lookup):
    prefix = get_prefix(component_code)
    if prefix in PREFIX_TO_COMMODITY_CODE:
        commodity_code = PREFIX_TO_COMMODITY_CODE[prefix]
        description = reference_lookup.get(commodity_code, "NOT FOUND IN REFERENCE FILE")
        return commodity_code, description, "OK"
    elif prefix in NEEDS_REVIEW_PREFIXES:
        return "NEEDS REVIEW", f"'{prefix}' has multiple possible categories", "REVIEW"
    else:
        return "NO MATCH", f"'{prefix}' not recognized", "REVIEW"


def build_excel_bytes(component_codes, reference_lookup):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    headers = ["Partno", "Component code in the drawing", "Mapping Code", "Description", "Status"]
    for col_num, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col_num, value=header).font = Font(bold=True)
    for row_num, code in enumerate(component_codes, start=2):
        commodity_code, description, status = map_component(code, reference_lookup)
        sheet.cell(row=row_num, column=1, value=row_num - 1)
        sheet.cell(row=row_num, column=2, value=code)
        sheet.cell(row=row_num, column=3, value=commodity_code)
        sheet.cell(row=row_num, column=4, value=description)
        sheet.cell(row=row_num, column=5, value=status)
    widths = {"A": 8, "B": 30, "C": 16, "D": 45, "E": 12}
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="CompMap | PCBA Mapping", page_icon="◈", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; }

    .hero {
        padding: 2.2rem 2rem;
        border-radius: 16px;
        background: linear-gradient(135deg, #6C5CE7 0%, #341f97 100%);
        margin-bottom: 1.8rem;
        box-shadow: 0 8px 32px rgba(108, 92, 231, 0.25);
    }
    .hero h1 {
        color: white; font-size: 2.1rem; font-weight: 700; margin: 0;
        letter-spacing: -0.02em;
    }
    .hero p {
        color: rgba(255,255,255,0.85); margin-top: 0.4rem; font-size: 0.95rem;
    }
    .badge {
        display: inline-block; background: rgba(255,255,255,0.15);
        color: white; padding: 3px 12px; border-radius: 20px;
        font-size: 0.7rem; font-family: 'JetBrains Mono', monospace;
        letter-spacing: 0.05em; margin-top: 0.6rem;
    }
    div[data-testid="stMetric"] {
        background: #1A1A2E; border: 1px solid #2D2D44;
        border-radius: 12px; padding: 1rem 1.2rem;
    }
    div[data-testid="stMetricValue"] {
        font-family: 'JetBrains Mono', monospace; color: #A29BFE;
    }
    .stDataFrame { border-radius: 12px; overflow: hidden; }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #6C5CE7; border-radius: 12px; padding: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>◈ CompMap</h1>
    <p>Upload a PCBA assembly drawing and commodity code table — get an instant, verified mapping report.</p>
    <span class="badge">v1.0 · PDF text extraction · zero OCR</span>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("**📄 Assembly Drawing**")
    pdf_file = st.file_uploader("PDF", type=["pdf"], label_visibility="collapsed")
with col2:
    st.markdown("**📋 Commodity Code Reference**")
    ref_file = st.file_uploader("Excel", type=["xlsx"], label_visibility="collapsed")

if pdf_file and ref_file:
    with st.spinner("Extracting and mapping components..."):
        codes = extract_component_codes(pdf_file.read())
        reference = load_reference_table(ref_file.read())

        rows = []
        for i, code in enumerate(codes, start=1):
            mcode, desc, status = map_component(code, reference)
            rows.append({
                "Partno": i, "Component code in the drawing": code,
                "Mapping Code": mcode, "Description": desc, "Status": status,
            })
        df = pd.DataFrame(rows)

    ok_count = int((df["Status"] == "OK").sum())
    review_count = int((df["Status"] == "REVIEW").sum())
    match_rate = round(100 * ok_count / len(df), 1) if len(df) else 0

    st.markdown("### Results")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total components", len(df))
    m2.metric("Confidently mapped", ok_count)
    m3.metric("Needs review", review_count)
    m4.metric("Match rate", f"{match_rate}%")

    # category breakdown chart
    ok_df = df[df["Status"] == "OK"].copy()
    if not ok_df.empty:
        cat_counts = ok_df["Mapping Code"].map(lambda c: CATEGORY_NAMES.get(c, c)).value_counts()
        fig = go.Figure(go.Bar(
            x=cat_counts.values, y=cat_counts.index, orientation='h',
            marker=dict(color='#6C5CE7', line=dict(width=0)),
        ))
        fig.update_layout(
            height=280, margin=dict(l=0, r=0, t=10, b=0),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#E8E8F0', family='Space Grotesk'),
            xaxis=dict(gridcolor='#2D2D44'), yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig, use_container_width=True)

    tab1, tab2 = st.tabs(["All components", "Needs review"])
    with tab1:
        st.dataframe(df, use_container_width=True, height=400)
    with tab2:
        st.dataframe(df[df["Status"] == "REVIEW"], use_container_width=True, height=300)

    excel_bytes = build_excel_bytes(codes, reference)
    st.download_button(
        "⬇ Download Excel Report", data=excel_bytes,
        file_name="Component_Mapping_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
else:
    st.info("Upload both files above to generate your report.")
