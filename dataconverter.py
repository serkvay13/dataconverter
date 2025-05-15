import os
import re
import pandas as pd
import numpy as np
import fitz  # PyMuPDF
import io
from PIL import Image
from langdetect import detect
import streamlit as st
from tempfile import NamedTemporaryFile
from openpyxl.styles import Alignment

# --- Configuration ---
LANGUAGES = {
    'en': 'en',
    'tr': 'tr',
    'fr': 'fr',
    'zh': 'ch_sim',
    'nl': 'nl'
}

KEYWORDS = {
    'chemicals': ['resin', 'chloride', 'fluoride', 'acid', 'soda', 'carbonate'],
    'additives': ['sugar', 'flavor', 'preservative', 'sweetener', 'color']
}

NACE_HS_MAP = {
    'PVC Resin': ('20.16', '3904.10'),
    'Calcium Chloride': ('20.13', '2827.20')
}

# --- OCR Reader with Caching ---
@st.cache_resource
def get_ocr_reader(lang_code='en'):
    import easyocr
    return easyocr.Reader([lang_code], gpu=False)

# --- PDF Text Extraction ---
def extract_text_from_pdf(pdf_path, lang_code='en'):
    reader = get_ocr_reader(lang_code)
    doc = fitz.open(pdf_path)
    text_blocks = []
    for page in doc:
        text = page.get_text()
        if not text.strip():
            try:
                pix = page.get_pixmap(dpi=150)
                img_bytes = pix.tobytes("png")
                image = Image.open(io.BytesIO(img_bytes))
                result = reader.readtext(np.array(image), detail=0, paragraph=True)
                text = "\n".join(result)
            except Exception as e:
                text = f"[OCR failed: {e}]"
        text_blocks.append(text)
    return "\n".join(text_blocks)

# --- Language Detection ---
def detect_language(text):
    try:
        return detect(text)
    except:
        return 'en'

# --- Text Parsing ---
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

# --- Product Filtering ---
def filter_products(products, category='chemicals'):
    keywords = KEYWORDS.get(category, [])
    return [prod for prod in products if any(kw.lower() in prod.lower() for kw in keywords)]

# --- NACE/HS Code Enrichment ---
def enrich_with_codes(products):
    enriched = []
    for p in products:
        nace, hs = NACE_HS_MAP.get(p, ('', ''))
        enriched.append((p, nace, hs))
    return enriched

# --- Excel Row Creation ---
def create_excel_row(company, contact, email, product_data):
    return {
        "Şirket İsmi": company,
        "İrtibat Bilgileri": contact,
        "E-Mail": email,
        "Ürünler": ", ".join([p[0] for p in product_data]),
        "NACE Kodları": ", ".join([p[1] for p in product_data if p[1]]),
        "HS Kodları": ", ".join([p[2] for p in product_data if p[2]])
    }

# --- PDF Processing Pipeline ---
def process_pdf_file(file_path, category='chemicals'):
    # İlk olarak tüm metni çıkar
    full_text = extract_text_from_pdf(file_path)
    # Dili tespit et ve uygun OCR reader ile tekrar dene
    lang = detect_language(full_text)
    lang_code = LANGUAGES.get(lang, 'en')
    if lang_code != 'en':
        full_text = extract_text_from_pdf(file_path, lang_code)
    company, contact, email, raw_products = parse_text(full_text)
    filtered_products = filter_products(raw_products, category=category)
    enriched = enrich_with_codes(filtered_products)
    return create_excel_row(company, contact, email, enriched)

# --- Streamlit App ---
def run_streamlit_app():
    st.title("Dataconverter - PDF to Excel Tool")
    uploaded_files = st.file_uploader("Upload one or more PDF files", accept_multiple_files=True, type="pdf")
    category = st.selectbox("Select product category for filtering", list(KEYWORDS.keys()))

    if uploaded_files and st.button("Process Files"):
        results = []
        for uploaded_file in uploaded_files:
            # Dosyayı geçici olarak kaydet
            with NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name
            try:
                row_data = process_pdf_file(tmp_path, category)
                results.append(row_data)
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

        if results:
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
                st.download_button(
                    label="Download Excel",
                    data=f,
                    file_name="converted_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == '__main__':
    run_streamlit_app()
