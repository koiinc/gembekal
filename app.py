import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import time
import re
import os
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient

# --- 1. AZURE SETUP ---
AZURE_ENDPOINT = "https://bekal-ocr.cognitiveservices.azure.com/"
# Using Streamlit Secrets to hide your API key safely!
AZURE_KEY = st.secrets["AZURE_KEY"]

# --- 2. PAGE SETUP ---
st.set_page_config(page_title="Document OCR Portal", layout="wide")
st.title("📄 Document Extraction & Template Portal")
st.write("Upload Borang Bekal forms, review the extracted data, and generate Word documents.")

# --- 3. BATCH UPLOAD ---
st.subheader("Step 1: Upload Documents")
uploaded_files = st.file_uploader("Upload Scanned Forms (PDF/Images)", accept_multiple_files=True)

if uploaded_files:
    st.info(f"📁 You have successfully uploaded {len(uploaded_files)} file(s).")

def extract_data_from_document(file):
    client = DocumentAnalysisClient(endpoint=AZURE_ENDPOINT, credential=AzureKeyCredential(AZURE_KEY))
    poller = client.begin_analyze_document("prebuilt-document", document=file.getvalue())
    result = poller.result()
    
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
    
    # 1. FIRST PASS: Try standard Key-Value Pairs for the easy/single-line fields
    if result.key_value_pairs:
        for kv_pair in result.key_value_pairs:
            if kv_pair.key and kv_pair.value:
                raw_key = kv_pair.key.content.upper()
                val_text = kv_pair.value.content.replace('\n', ' ').strip()
                
                if "EMEL" in raw_key: extracted_data["EMEL"] = val_text
                elif "TEL" in raw_key: extracted_data["NO_TEL"] = val_text
                elif "TARIKH" in raw_key and not extracted_data["TARIKH"]: extracted_data["TARIKH"] = val_text

    # 2. SECOND PASS: Advanced Boundary Checking for complex/multi-line table fields
    all_lines = [line.content.strip() for page in result.pages for line in page.lines]

    # --- NAMA PEMBEKAL BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "NAMA PEMBEKAL" in line_text.upper():
            collected_name = []
            
            # Capture anything on the exact same line
            remainder = re.sub(r'(?i)NAMA PEMBEKAL\s*:?\s*', '', line_text).strip()
            if remainder:
                collected_name.append(remainder)
                
            # Look at the next few lines for the second line of the name
            for j in range(i + 1, min(i + 5, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                
                # THE STOPPING WALL (Looks for the EMEL cell to the right or NAMA SYARIKAT below)
                if "EMEL" in next_line_upper or "NAMA SYARIKAT" in next_line_upper:
                    break
                    
                collected_name.append(next_line)
                
            if collected_name:
                extracted_data["NAMA_PEMBEKAL"] = " ".join(collected_name).strip()
            break

    # --- NAMA SYARIKAT BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "NAMA SYARIKAT" in line_text.upper(): 
            collected_company = []
            
            # Capture anything on the exact same line
            remainder = re.sub(r'(?i)NAMA SYARIKAT\s*:?\s*', '', line_text).strip()
            if remainder:
                collected_company.append(remainder)
            
            # Read downwards to get the rest of the company name
            for j in range(i + 1, min(i + 5, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                
                # THE STOPPING WALLS (Fields to the right or directly below)
                if "ALAMAT" in next_line_upper or "TARIKH" in next_line_upper or "PENSIJILAN" in next_line_upper:
                    break
                
                collected_company.append(next_line)
            
            if collected_company:
                extracted_data["NAMA_SYARIKAT"] = " ".join(collected_company).strip()
            break

    # --- ALAMAT PREMIS BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "ALAMAT PREMIS" in line_text.upper():
            collected_alamat = []
            
            # Capture anything on the exact same line
            remainder = re.sub(r'(?i)ALAMAT PREMIS\s*:?\s*', '', line_text).strip()
            if remainder:
                collected_alamat.append(remainder)
                
            # Look at the next few lines for the rest of the address
            for j in range(i + 1, min(i + 8, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                
                # THE STOPPING WALL (Looks for the fields in the row directly below it)
                if "TARIKH" in next_line_upper or "PENSIJILAN" in next_line_upper or "BESS" in next_line_upper:
                    break
                    
                collected_alamat.append(next_line)
                
            if collected_alamat:
                extracted_data["ALAMAT_PREMIS"] = " ".join(collected_alamat).strip()
            break

    # --- TAJUK PEROLEHAN BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "TAJUK PEROLEHAN" in line_text.upper():
            collected_title = []
            
            # Capture text on the exact same line
            remainder = re.sub(r'(?i)TAJUK PEROLEHAN\s*:?\s*', '', line_text).strip()
            if remainder:
                collected_title.append(remainder)
                
            # Look at the next few lines for the rest of the text
            for j in range(i + 1, min(i + 5, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                
                # THE STOPPING WALL (Looks for the field directly below it)
                if "TARIKH PERMOHONAN" in next_line_upper or "DITERIMA" in next_line_upper:
                    break 
                
                collected_title.append(next_line)
            
            if collected_title:
                extracted_data["TAJUK_PEROLEHAN"] = " ".join(collected_title).strip()
            break 

    return extracted_data

if uploaded_files:
    if st.button("🔍 Run OCR Extraction"):
        with st.spinner("Extracting data with Azure... Please wait."):
            extracted_records = []
            for file in uploaded_files:
                try:
                    data = extract_data_from_document(file)
                    extracted_records.append(data)
                except Exception as e:
                    st.error(f"Failed to process {file.name}: {e}")
                
                # Pause for 1 second between files to respect Azure Free Tier limits
                time.sleep(1) 
            
            if extracted_records:
                df = pd.DataFrame(extracted_records)
                
                # --- DATA CLEANING STEP ---
                if 'NAMA_SYARIKAT' in df.columns:
                    # Remove IC/Registration Numbers from the start of NAMA_SYARIKAT
                    df['NAMA_SYARIKAT'] = df['NAMA_SYARIKAT'].astype(str).str.replace(r'^\d{6}-\d{2}-\d{4}\s*', '', regex=True)

                if 'TARIKH' in df.columns:
                    # Clear out the Date if it grabbed the boilerplate text
                    df['TARIKH'] = df['TARIKH'].apply(lambda x: "" if "BELUM LENGKAP" in str(x).upper() else x)
                
                st.session_state['ocr_data'] = df
                st.success("Extraction Complete!")

# --- 4. REVIEW AND EDIT TABLE ---
if 'ocr_data' in st.session_state:
    st.subheader("Step 2: Review and Edit Data")
    st.write("Click on any cell below to fix typos or manually paste missing data before generating the Word documents.")
    
    edited_df = st.data_editor(st.session_state['ocr_data'], num_rows="dynamic", use_container_width=True)

    # --- 5. GENERATE WORD DOCS ---
    st.subheader("Step 3: Generate Output")
    
    template_file = "Borang Bekal Template.docx" 
    
    if st.button("📝 Generate Word Documents"):
        if not os.path.exists(template_file):
            st.error(f"Template file '{template_file}' not found. Please ensure it is in the same directory as this script.")
        else:
            with st.spinner("Creating documents..."):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for index, row in edited_df.iterrows():
                        doc = DocxTemplate(template_file)
                        context = row.dropna().to_dict()
                        doc.render(context)
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        safe_name = str(row.get('NAMA_PEMBEKAL', 'TiadaNama')).replace("/", "-")
                        if not safe_name or safe_name.lower() == 'nan':
                            safe_name = "TiadaNama"
                            
                        final_filename = f"Completed_{index + 1}_{safe_name}.docx"
                        zip_file.writestr(final_filename, doc_io.getvalue())
                
                st.download_button(
                    label="⬇️ Download All Word Documents (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Completed_Documents.zip",
                    mime="application/zip"
                )
