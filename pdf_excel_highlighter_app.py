
import streamlit as st
import pandas as pd
from openpyxl import Workbook
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
v.1.20
""")

# Upload files
uploaded_pdfs = st.file_uploader("Ladda upp PDF-filer", type=["pdf"], accept_multiple_files=True)
uploaded_reference = st.file_uploader("Ladda upp ritningsförteckning", type=["xlsx", "pdf"])

# Start button
start_processing = st.button("Starta jämförelse")

# Clean text function
def clean_text(text):
    text = str(text).strip().lower()
    return re.sub(r'\.(pdf|docx?|xlsx?|txt|jpg|png|csv)$', '', text)

# Flexible regex for drawing names
drawing_pattern = re.compile(r'^(?=.*\d)[a-z0-9]+([-_][a-z0-9]+){2,}$', re.IGNORECASE)

# Generic words to exclude
exclude_terms = {"plan", "del", "sektion", "fasad", "1:50", "1:100"}

if start_processing and uploaded_pdfs and uploaded_reference:
    with st.spinner("Bearbetar filer..."):
        try:
            progress = st.progress(0)
            total_steps = 6

            # Step 1: Extract and clean PDF filenames
            pdf_names = [pdf.name for pdf in uploaded_pdfs]
            cleaned_pdf_names = [clean_text(name) for name in pdf_names]
            progress.progress(1 / total_steps)

            # Step 2: Read and clean reference text
            reference_texts = []

            if uploaded_reference.name.lower().endswith(".xlsx"):
                excel_bytes = BytesIO(uploaded_reference.read())
                df_ref = pd.read_excel(excel_bytes, header=None, engine="openpyxl")
                reference_texts = df_ref.astype(str).stack().map(clean_text).tolist()

            elif uploaded_reference.name.lower().endswith(".pdf"):
                pdf_bytes = BytesIO(uploaded_reference.read())
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(pdf_bytes.read())
                    temp_pdf.flush()
                    with pdfplumber.open(temp_pdf.name) as pdf:
                        tables_found = False
                        for page in pdf.pages:
                            tables = page.extract_tables()
                            if tables:
                                tables_found = True
                                for table in tables:
                                    for row in table:
                                        for cell in row:
                                            if cell:
                                                reference_texts.append(clean_text(cell))
                        if not tables_found:
                            for page in pdf.pages:
                                page_text = page.extract_text()
                                if page_text:
                                    lines = [line.strip() for line in page_text.splitlines() if line.strip()]
                                    for line in lines:
                                        reference_texts.append(clean_text(line))

            progress.progress(2 / total_steps)

            # Step 3: Filter reference texts using regex and exclude generic terms
            filtered_reference_texts = [
                ref for ref in reference_texts
                if drawing_pattern.match(ref)
                and not any(term in ref for term in exclude_terms)
                and not re.match(r'^202\d', ref)  # Exclude anything starting with 2020, 2021, etc.
            ]

            # Step 4: Create Excel with match status and highlight matches
            wb = Workbook()
            ws = wb.active
            ws.title = "Ritningsförteckning"

            # Header
            ws.append(["Referensnamn", "Matchstatus"])

            # Highlight style
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

            for ref in filtered_reference_texts:
                match_status = "Matchad" if ref in cleaned_pdf_names else "Ej matchad"
                row = [ref, match_status]
                ws.append(row)
                if match_status == "Matchad":
                    for cell in ws[ws.max_row]:
                        cell.fill = fill

            progress.progress(3 / total_steps)

            # Step 5: Save highlighted Excel file
            result_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            wb.save(result_file.name)
            progress.progress(4 / total_steps)

            # Step 6: Save unmatched PDF names
            unmatched_cleaned = [name for name in cleaned_pdf_names if name not in filtered_reference_texts]
            df_unmatched = pd.DataFrame(unmatched_cleaned, columns=["Omatchade PDF-namn"])
            unmatched_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df_unmatched.to_excel(unmatched_file.name, index=False)
            progress.progress(1.0)

            st.success("Bearbetning klar!")
            st.download_button("Ladda ner ritningsförteckning med markering", data=open(result_file.name, "rb").read(), file_name="ritningsförteckning_markering.xlsx")
            st.download_button("Ladda ner lista på omatchade PDFer", data=open(unmatched_file.name, "rb").read(), file_name="omatchade_ritningar.xlsx")

        except Exception as e:
            st.error(f"Ett fel uppstod: {e}")
