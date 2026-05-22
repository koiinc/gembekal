import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import time
import re # <--- ADD THIS NEW LINE HERE
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

# Improvement 5: Tell the user exactly how many files they just uploaded
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
    
    # 1. FIRST PASS: Try standard Key-Value Pairs for the easy fields
    if result.key_value_pairs:
        for kv_pair in result.key_value_pairs:
            if kv_pair.key and kv_pair.value:
                raw_key = kv_pair.key.content.upper()
                val_text = kv_pair.value.content.replace('\n', ' ').strip()
                
                if "NAMA PEMBEKAL" in raw_key: extracted_data["NAMA_PEMBEKAL"] = val_text
                elif "EMEL" in raw_key: extracted_data["EMEL"] = val_text
                elif "TEL" in raw_key: extracted_data["NO_TEL"] = val_text
                elif "ALAMAT" in raw_key and "PREMIS" in raw_key: extracted_data["ALAMAT_PREMIS"] = val_text
                elif "TARIKH" in raw_key and not extracted_data["TARIKH"]: extracted_data["TARIKH"] = val_text

    # 2. SECOND PASS: Advanced Boundary Checking for Syarikat and Tajuk
    all_lines = []
    for page in result.pages:
        for line in page.lines:
            all_lines.append(line.content.strip())
            
    # --- TAJUK PEROLEHAN BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "TAJUK" in line_text.upper() or "PEROLEHAN" in line_text.upper():
            collected_title = []
            # Look at the next few lines (up to 10 lines max)
            for j in range(i + 1, min(i + 10, len(all_lines))):
                next_line = all_lines[j].upper()
                
                # THE STOPPING WALL
                if "TARIKH PERMOHONAN LENGKAP DITERIMA" in next_line:
                    break 
                
                collected_title.append(all_lines[j])
            
            if collected_title:
                extracted_data["TAJUK_PEROLEHAN"] = " ".join(collected_title).strip()
            break # Stop searching the document once we found it

    # --- NAMA SYARIKAT BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        # Find the label
        if "SYARIKAT" in line_text.upper(): 
            collected_company = []
            # Look at the next few lines (up to 5 lines max)
            for j in range(i + 1, min(i + 5, len(all_lines))):
                next_line = all_lines[j].upper()
                
                # THE STOPPING WALLS (Label below, or label to the right)
                if "TARIKH" in next_line or "ALAMAT" in next_line:
                    break
                
                collected_company.append(all_lines[j])
            
            if collected_company:
                raw_company = " ".join(collected_company).strip()
                # Clean up IC numbers using regex (removes XXXX-XX-XXXX format)
                clean_company = re.sub(r'\d{6}-\d{2}-\d{4}\s*', '', raw_company)
                extracted_data["NAMA_SYARIKAT"] = clean_company.strip()
            break # Stop searching the document once we found it

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
            
            # Create the DataFrame
            df = pd.DataFrame(extracted_records)
            
            # --- DATA CLEANING STEP ---
            # Improvement 3: Remove IC/Registration Numbers from the start of NAMA_SYARIKAT
            if 'NAMA_SYARIKAT' in df.columns:
                df['NAMA_SYARIKAT'] = df['NAMA_SYARIKAT'].astype(str).str.replace(r'^\d{6}-\d{2}-\d{4}\s*', '', regex=True)

            # Improvement 4: Clear out the Date if it grabbed the "BELUM LENGKAP" boilerplate text
            if 'TARIKH' in df.columns:
                df['TARIKH'] = df['TARIKH'].apply(lambda x: "" if "BELUM LENGKAP" in str(x) else x)
            # --------------------------
            
            st.session_state['ocr_data'] = df
            st.success("Extraction Complete!")

# --- 4. REVIEW AND EDIT TABLE ---
if 'ocr_data' in st.session_state:
    st.subheader("Step 2: Review and Edit Data")
    st.write("Click on any cell below to fix typos or manually paste missing data (like TAJUK_PEROLEHAN) before generating the Word documents.")
    
    edited_df = st.data_editor(st.session_state['ocr_data'], num_rows="dynamic", use_container_width=True)

    # --- 5. GENERATE WORD DOCS ---
    st.subheader("Step 3: Generate Output")
    
    # Improvement 2: Hardcode the template file instead of asking for upload!
    # IMPORTANT: Ensure your Word file in GitHub is exactly named "Borang Bekal Template.docx"
    template_file = "Borang Bekal Template.docx" 
    
    if st.button("📝 Generate Word Documents"):
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
                    safe_name = str(row['NAMA_PEMBEKAL']).replace("/", "-") if row['NAMA_PEMBEKAL'] else "TiadaNama"
                    
                    # Improvement 1: Add row index + 1 so files NEVER overwrite each other in the ZIP
                    final_filename = f"Completed_{index + 1}_{safe_name}.docx"
                    
                    zip_file.writestr(final_filename, doc_io.getvalue())
            
            st.download_button(
                label="⬇️ Download All Word Documents (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Completed_Documents.zip",
                mime="application/zip"
            )
