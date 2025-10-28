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

            # Step 1: Extract and clean PDF filenames
            pdf_names = [pdf.name for pdf in uploaded_pdfs]
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

            if uploaded_reference.name.lower().endswith(".xlsx"):
                excel_bytes = BytesIO(uploaded_reference.read())
                df_ref = pd.read_excel(excel_bytes, header=None, engine="openpyxl")
                reference_texts = set(df_ref.astype(str).stack().map(clean_text).unique())

            elif uploaded_reference.name.lower().endswith(".pdf"):
                pdf_bytes = BytesIO(uploaded_reference.read())
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(pdf_bytes.read())
                    temp_pdf.flush()
                    with pdfplumber.open(temp_pdf.name) as pdf:
                        text = ""
                        for page in pdf.pages:
                            page_text = page.extract_text()
                            if page_text:
                                text += page_text + "\n"
                lines = [line.strip() for line in text.splitlines() if line.strip()]
                reference_texts = set([clean_text(line) for line in lines])

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
