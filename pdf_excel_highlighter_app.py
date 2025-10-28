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
OBS: Ritningsförteckning fungerar bara som Excel just nu.  
v.1.12
""")

# Initialize session state
if "uploaded_pdfs" not in st.session_state:
    st.session_state.uploaded_pdfs = None
if "uploaded_reference" not in st.session_state:
    st.session_state.uploaded_reference = None

# File uploaders
st.session_state.uploaded_pdfs = st.file_uploader("Ladda upp PDF-filer", type=["pdf"], accept_multiple_files=True)
st.session_state.uploaded_reference = st.file_uploader("Ladda upp ritningsförteckning", type=["xlsx", "pdf"])

# Reset button
if st.button("Återställ filer"):
    st.session_state.uploaded_pdfs = None
    st.session_state.uploaded_reference = None
    st.experimental_rerun()

# Start processing button
start_processing = st.button("Starta jämförelse")

if start_processing and st.session_state.uploaded_pdfs and st.session_state.uploaded_reference:
    with st.spinner("Bearbetar filer..."):
        try:
            progress = st.progress(0)
            total_steps = 6

            # Step 1: Extract and clean PDF filenames
            pdf_names = [pdf.name for pdf in st.session_state.uploaded_pdfs]
            cleaned_pdf_names = [re.sub(r'\.(pdf)$', '', name.strip().lower()) for name in pdf_names]
            df_pdf_list = pd.DataFrame(pdf_names, columns=['File Name'])
            progress.progress(1 / total_steps)

            # Step 2: Save PDF list to temporary Excel file
            temp_file2 = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df_pdf_list.to_excel(temp_file2.name, index=False)
            progress.progress(2 / total_steps)

            # Step 3: Read and clean reference text
            def clean_text(text):
                text = str(text).strip().lower()
                return re.sub(r'\.(pdf|docx?|xlsx?|txt|jpg|png|csv)$', '', text)

            if st.session_state.uploaded_reference.name.lower().endswith(".xlsx"):
                excel_bytes = BytesIO(st.session_state.uploaded_reference.read())
                df_ref = pd.read_excel(excel_bytes, header=None, engine="openpyxl")
                reference_texts = set(df_ref.astype(str).stack().map(clean_text).unique())

                        
            elif st.session_state.uploaded_reference.name.lower().endswith(".pdf"):
                pdf_bytes = BytesIO(st.session_state.uploaded_reference.read())
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(pdf_bytes.read())
                    temp_pdf.flush()
                    with pdfplumber.open(temp_pdf.name) as pdf:
                        reference_texts = set()
                        tables_found = False
            
                        for page in pdf.pages:
                            tables = page.extract_tables()
                            if tables:
                                tables_found = True
                                for table in tables:
                                    for row in table:
                                        for cell in row:
                                            if cell:
                                                reference_texts.add(clean_text(cell))

            # Fallback to text extraction if no tables were found
            if not tables_found:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        lines = [line.strip() for line in page_text.splitlines() if line.strip()]
                        for line in lines:
                            reference_texts.add(clean_text(line))


            progress.progress(3 / total_steps)

            # Step 4: Highlight matches in Excel file
            wb = load_workbook(temp_file2.name)
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value:
                            cleaned_value = clean_text(cell.value)
                            if cleaned_value in reference_texts:
                                cell.fill = fill

            progress.progress(4 / total_steps)

            # Step 5: Save highlighted Excel file
            result_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            wb.save(result_file.name)
            progress.progress(5 / total_steps)

            # Step 6: Save unmatched PDF names
            unmatched_cleaned = [name for name in cleaned_pdf_names if name not in reference_texts]
            df_unmatched = pd.DataFrame(unmatched_cleaned, columns=["Omatchade PDF-namn"])
            unmatched_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df_unmatched.to_excel(unmatched_file.name, index=False)
            progress.progress(1.0)

            st.success("Bearbetning klar!")
            st.download_button("Ladda ner ritningsförteckning med markering", data=open(result_file.name, "rb").read(), file_name="ritningsförteckning_markering.xlsx")
            st.download_button("Ladda ner lista på omatchade PDFer", data=open(unmatched_file.name, "rb").read(), file_name="omatchade_ritningar.xlsx")

        except Exception as e:
            st.error(f"Ett fel uppstod: {e}")
