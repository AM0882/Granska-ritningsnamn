import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import tempfile
import pdfplumber
from io import BytesIO

st.title("Jämför ritningsförteckning med PDF-filer")

st.markdown("""
Ladda upp ritningar och ritningsförteckning och jämför.  
I resultatet fås en ritningsförteckning där alla ritningar som finns med som PDF är gulmarkerade,  
samt en lista på de ritningar som är med som PDF men inte finns i förteckning.
v.2
""")

# Step 1: Upload multiple PDF files
uploaded_pdfs = st.file_uploader("Ladda upp PDF-filer", type=["pdf"], accept_multiple_files=True)

# Step 2: Upload the reference file (Excel or PDF)
uploaded_reference = st.file_uploader("Ladda upp ritningsförteckning", type=["xlsx", "pdf"])

if uploaded_pdfs and uploaded_reference:
    with st.spinner("Bearbetar filer..."):
        try:
            progress = st.progress(0)
            total_steps = 6
            # Improved cleaning function
            def clean_text(text):
                text = str(text).strip().lower()
                text = re.sub(r'\.(pdf|docx?|xlsx?|txt|jpg|png|csv)$', '', text)
                text = re.sub(r'\s+', '', text)  # Remove all whitespace
                text = re.sub(r'[^a-z0-9]', '', text)  # Remove non-alphanumeric characters
                return text

            # Step 1: Extract and clean PDF filenames
            pdf_names = [pdf.name for pdf in uploaded_pdfs]
            cleaned_pdf_names = [clean_text(name) for name in pdf_names]
            df_pdf_list = pd.DataFrame(pdf_names, columns=['File Name'])
            progress.progress(1 / total_steps)

            # Step 2: Save PDF list to temporary Excel file
            temp_file2 = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df_pdf_list.to_excel(temp_file2.name, index=False)
            progress.progress(2 / total_steps)

            # Step 3: Read and clean reference text
            if uploaded_reference.name.lower().endswith(".xlsx"):
                excel_bytes = BytesIO(uploaded_reference.read())
                df_ref = pd.read_excel(excel_bytes, header=None, engine="openpyxl")
                reference_texts = set(df_ref.astype(str).stack().map(clean_text).unique())

            elif uploaded_reference.name.lower().endswith(".pdf"):
                pdf_bytes = BytesIO(uploaded_reference.read())
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as
