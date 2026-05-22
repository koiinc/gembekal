import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import time
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient

# --- 1. AZURE SETUP ---
AZURE_ENDPOINT = "https://bekal-ocr.cognitiveservices.azure.com/" 
AZURE_KEY = st.secrets["AZURE_KEY"]

# --- 2. PAGE SETUP ---
st.set_page_config(page_title="Document OCR Portal", layout="wide")
st.title("📄 Document Extraction & Template Portal")
st.write("Upload Borang Bekal forms, review the extracted data, and generate Word documents.")

# --- 3. BATCH UPLOAD ---
st.subheader("Step 1: Upload Documents")
uploaded_files = st.file_uploader("Upload Scanned Forms (PDF/Images)", accept_multiple_files=True)

def extract_data_from_document(file):
    # Connect to your specific Azure account
    client = DocumentAnalysisClient(endpoint=AZURE_ENDPOINT, credential=AzureKeyCredential(AZURE_KEY))
    
    # Send the file to Azure
    poller = client.begin_analyze_document("prebuilt-document", document=file.getvalue())
    result = poller.result()
    
    # DEFINE ONLY THE COLUMNS YOU WANT (Matches your Word tags exactly)
    extracted_data = {
        "Filename": file.name,
        "NAMA_PEMBEKAL": "",
        "EMEL": "",
        "NO_TEL": "",
        "NAMA_SYARIKAT": "",
        "ALAMAT_PREMIS": "",
        "TARIKH": "",
        "TAJUK_PEROLEHAN": ""
    }
    
    # LOOP THROUGH AND FILTER
    if result.key_value_pairs:
        for kv_pair in result.key_value_pairs:
            if kv_pair.key and kv_pair.value:
                raw_key = kv_pair.key.content.upper()
                
                # THIS FIXES THE 2-LINE PROBLEM: Replaces newlines with a normal space
                val_text = kv_pair.value.content.replace('\n', ' ').strip()
                
                # MATCH TO YOUR TEMPLATE TAGS
                if "NAMA PEMBEKAL" in raw_key:
                    extracted_data["NAMA_PEMBEKAL"] = val_text
                elif "EMEL" in raw_key:
                    extracted_data["EMEL"] = val_text
                elif "TEL" in raw_key:
                    extracted_data["NO_TEL"] = val_text
                elif "SYARIKAT" in raw_key:
                    extracted_data["NAMA_SYARIKAT"] = val_text
                elif "ALAMAT" in raw_key and "PREMIS" in raw_key:
                    extracted_data["ALAMAT_PREMIS"] = val_text
                elif "TARIKH" in raw_key and not extracted_data["TARIKH"]: 
                    extracted_data["TARIKH"] = val_text
                elif "TAJUK" in raw_key or "PEROLEHAN" in raw_key:
                    extracted_data["TAJUK_PEROLEHAN"] = val_text
                
    return extracted_data

if uploaded_files:
    if st.button("🔍 Run OCR Extraction"):
        with st.spinner("Extracting data with Azure... Please wait."):
            extracted_records = []
            for file in uploaded_files:
                data = extract_data_from_document(file)
                extracted_records.append(data)
                # Pause for 1 second between files so we don't overwhelm Azure's Free Tier
                time.sleep(1) 
            
            st.session_state['ocr_data'] = pd.DataFrame(extracted_records)
            st.success("Extraction Complete!")

# --- 4. REVIEW AND EDIT TABLE ---
if 'ocr_data' in st.session_state:
    st.subheader("Step 2: Review and Edit Data")
    st.write("Click on any cell below to fix typos before generating the Word documents.")
    
    edited_df = st.data_editor(st.session_state['ocr_data'], num_rows="dynamic", use_container_width=True)

    # --- 5. GENERATE WORD DOCS ---
    st.subheader("Step 3: Generate Output")
    template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])
    
    if template_file and st.button("📝 Generate Word Documents"):
        with st.spinner("Creating documents..."):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for index, row in edited_df.iterrows():
                    doc = DocxTemplate(template_file)
                    
                    # Convert the row to a dictionary, dropping any empty values
                    context = row.dropna().to_dict()
                    doc.render(context)
                    
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)
                    
                    # Create a safe file name based on the NAMA_PEMBEKAL
                    safe_name = str(row['NAMA_PEMBEKAL']).replace("/", "-") if row['NAMA_PEMBEKAL'] else f"Document_{index}"
                    zip_file.writestr(f"Completed_{safe_name}.docx", doc_io.getvalue())
            
            st.download_button(
                label="⬇️ Download All Word Documents (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Completed_Documents.zip",
                mime="application/zip"
            )
