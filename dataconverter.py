# dataconverter: Local OCR PDF to Excel Converter Tool
# Features:
# - Multi-language OCR
# - Structured Excel export
# - Product filtering
# - NACE/HS Code prediction
# - Optional Google Sheets integration
# - Multiple PDF file processing
# - Streamlit-based GUI
# dataconverter: Local OCR PDF to Excel Converter Tool
# Features:
# - Multi-language OCR
# - Structured Excel export
# - Product filtering
# - NACE/HS Code prediction
# - Optional Google Sheets integration
# - Multiple PDF file processing
# - Streamlit-based GUI

import os
import re
import pytesseract
if os.name == 'nt':
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
from pdf2image import convert_from_path
import pandas as pd
from langdetect import detect
from fuzzywuzzy import process
import streamlit as st
from tempfile import NamedTemporaryFile
from openpyxl.styles import Alignment

# --- Poppler Path Configuration for Windows ---

# --- Configuration ---
LANGUAGES = {
    'en': 'eng',
    'tr': 'tur',
    'fr': 'fra',
    'zh': 'chi_sim',
    'nl': 'nld'
}

KEYWORDS = {
    'chemicals': ['resin', 'chloride', 'fluoride', 'acid', 'soda', 'carbonate'],
    'additives': ['sugar', 'flavor', 'preservative', 'sweetener', 'color']
}

NACE_HS_MAP = {
    'PVC Resin': ('20.16', '3904.10'),
    'Calcium Chloride': ('20.13', '2827.20')
}

import fitz  # PyMuPDF
from PIL import Image
import io

def extract_text_from_pdf(pdf_path, lang_code='eng'):
    doc = fitz.open(pdf_path)
    return "\n".join([page.get_text() for page in doc])







def detect_language(text):
    try:
        return detect(text)
    except:
        return 'en'

def parse_text(text):
    lines = text.splitlines()
    company_name, contact_info, email, products = '', '', '', []
    for line in lines:
        line_lower = line.lower()
        if 'tel' in line_lower or 'adres' in line_lower or 'www' in line_lower or 'http' in line_lower or '@' in line_lower:
            contact_info += line.strip() + '\n'
        if '@' in line:
            found = re.findall(r'[\w\.-]+@[\w\.-]+', line)
            if found:
                email = found[0]
        elif any(word in line_lower for word in ['group', 'company', 'co.', 'inc.', 'ltd']):
            company_name = line.strip()
        elif len(line.split()) <= 5 and any(c.isalpha() for c in line):
            products.append(line.strip())
    return company_name, contact_info.strip(), email, list(set(products))

def filter_products(products, category='chemicals'):
    keywords = KEYWORDS.get(category, [])
    return [prod for prod in products if any(kw.lower() in prod.lower() for kw in keywords)]

def enrich_with_codes(products):
    enriched = []
    for p in products:
        nace, hs = NACE_HS_MAP.get(p, ('', ''))
        enriched.append((p, nace, hs))
    return enriched

def create_excel_row(company, contact, email, product_data):
    return {
        "Şirket İsmi": company,
        "İrtibat Bilgileri": contact,
        "E-Mail": email,
        "Ürünler": ", ".join([p[0] for p in product_data]),
        "NACE Kodları": ", ".join([p[1] for p in product_data if p[1]]),
        "HS Kodları": ", ".join([p[2] for p in product_data if p[2]])
    }

def process_pdf_file(file_path, category='chemicals'):
    full_text = extract_text_from_pdf(file_path)
    lang = detect_language(full_text)
    lang_code = LANGUAGES.get(lang, 'eng')
    full_text = extract_text_from_pdf(file_path, lang_code)
    company, contact, email, raw_products = parse_text(full_text)
    filtered_products = filter_products(raw_products, category=category)
    enriched = enrich_with_codes(filtered_products)
    return create_excel_row(company, contact, email, enriched)

def run_streamlit_app():
    st.title("Dataconverter - PDF to Excel Tool")
    uploaded_files = st.file_uploader("Upload one or more PDF files", accept_multiple_files=True, type="pdf")
    category = st.selectbox("Select product category for filtering", list(KEYWORDS.keys()))

    if uploaded_files and st.button("Process Files"):
        results = []
        for uploaded_file in uploaded_files:
            with NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name
            row_data = process_pdf_file(tmp_path, category)
            results.append(row_data)
            os.unlink(tmp_path)

        df = pd.DataFrame(results)
        st.dataframe(df)

        output_path = "converted_output.xlsx"
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.book.active
            col_index = df.columns.get_loc("Ürünler") + 1
            col_letter = chr(64 + col_index)
            worksheet.column_dimensions[col_letter].width = 50
            for row in worksheet.iter_rows(min_row=2, max_row=len(df)+1, min_col=col_index, max_col=col_index):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True)

        with open(output_path, "rb") as f:
            st.download_button(label="Download Excel", data=f, file_name="converted_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    run_streamlit_app()
