import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage # Updated import
from docx.shared import Mm # Import for image sizing
import io
import zipfile
import time
import re
import os
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient

# --- [ADD NEW IMPORTS HERE] ---
from PIL import Image, ImageOps, ImageFilter
# --------------------------------

# --- 1. AZURE SETUP ---
AZURE_ENDPOINT = "https://bekal-ocr.cognitiveservices.azure.com/"
# Using Streamlit Secrets to hide your API key safely!
AZURE_KEY = st.secrets["AZURE_KEY"]

# --- [NANDATANGAN - IMAGE PROCESSING FUNCTION] ---
def process_signature(uploaded_file):
    """
    Takes an uploaded image file (bytes), removes the white background, 
    makes it semi-transparent, and sharpens it.
    """
    try:
        # Open image and convert to RGBA to allow transparency
        img = Image.open(uploaded_file).convert("RGBA")
        datas = img.get_data()

        newData = []
        # Background Removal Logic (Pillow)
        # We replace pixels that are 'white-ish' (above 230 in RGB channels) 
        # with a completely transparent pixel (alpha = 0)
        tolerance = 230
        for item in datas:
            if item[0] > tolerance and item[1] > tolerance and item[2] > tolerance:
                # Replace white pixel with transparent pixel
                newData.append((255, 255, 255, 0))
            else:
                # Translucence: Make non-white pixels (the ink) semi-transparent (180 out of 255)
                # You can reduce 180 to make it more see-through, increase it to make it bolder.
                newData.append((item[0], item[1], item[2], 180)) 
        
        img.putdata(newData)
        
        # Sharpening Filter
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.SHARPEN) # Sharpen twice for boldness

        # Save to a byte stream
        img_bytes = io.BytesIO()
        img.save(img_bytes, format='PNG') # Must be PNG to preserve transparency
        img_bytes.seek(0)
        return img_bytes
    except Exception as e:
        st.error(f"Error processing signature image: {e}")
        return None
# ---------------------------------------------------

# --- 2. PAGE SETUP ---
st.set_page_config(page_title="Document OCR Portal", layout="wide")
st.title("📄 Document Extraction & Template Portal")
st.write("Upload Borang Bekal forms, review the extracted data, and generate Word documents.")

# --- SESSION STATE INITIALIZATION ---
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = "init"

# This is our running memory for the ZIP folder number!
if "zip_counter" not in st.session_state:
    st.session_state.zip_counter = 1

# --- 3. BATCH UPLOAD ---
st.subheader("Step 1: Upload Documents")
uploaded_files = st.file_uploader(
    "Upload Scanned Forms (PDF/Images)", 
    accept_multiple_files=True,
    key=st.session_state.uploader_key
)

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
    
    # 1. FIRST PASS: Try standard Key-Value Pairs
    if result.key_value_pairs:
        for kv_pair in result.key_value_pairs:
            if kv_pair.key and kv_pair.value:
                raw_key = kv_pair.key.content.upper()
                val_text = kv_pair.value.content.replace('\n', ' ').strip()
                
                if "EMEL" in raw_key: extracted_data["EMEL"] = val_text
                elif "TEL" in raw_key: extracted_data["NO_TEL"] = val_text
                elif "TARIKH" in raw_key and not extracted_data["TARIKH"]: extracted_data["TARIKH"] = val_text

    # 2. SECOND PASS: Advanced Boundary Checking
    all_lines = [line.content.strip() for page in result.pages for line in page.lines]

    # --- NAMA PEMBEKAL BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "NAMA PEMBEKAL" in line_text.upper():
            collected_name = []
            remainder = re.sub(r'(?i)NAMA PEMBEKAL\s*:?\s*', '', line_text).strip()
            if remainder: collected_name.append(remainder)
                
            for j in range(i + 1, min(i + 5, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                if "EMEL" in next_line_upper or "NAMA SYARIKAT" in next_line_upper: break
                collected_name.append(next_line)
                
            if collected_name: extracted_data["NAMA_PEMBEKAL"] = " ".join(collected_name).strip()
            break

    # --- NAMA SYARIKAT BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "NAMA SYARIKAT" in line_text.upper(): 
            collected_company = []
            remainder = re.sub(r'(?i)NAMA SYARIKAT\s*:?\s*', '', line_text).strip()
            if remainder: collected_company.append(remainder)
            
            for j in range(i + 1, min(i + 5, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                if "ALAMAT" in next_line_upper or "TARIKH" in next_line_upper or "PENSIJILAN" in next_line_upper: break
                collected_company.append(next_line)
            
            if collected_company: extracted_data["NAMA_SYARIKAT"] = " ".join(collected_company).strip()
            break

    # --- ALAMAT PREMIS BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "ALAMAT PREMIS" in line_text.upper():
            collected_alamat = []
            remainder = re.sub(r'(?i)ALAMAT PREMIS\s*:?\s*', '', line_text).strip()
            if remainder: collected_alamat.append(remainder)
                
            for j in range(i + 1, min(i + 8, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                if "TARIKH" in next_line_upper or "PENSIJILAN" in next_line_upper or "BESS" in next_line_upper: break
                collected_alamat.append(next_line)
                
            if collected_alamat: extracted_data["ALAMAT_PREMIS"] = " ".join(collected_alamat).strip()
            break

    # --- TAJUK PEROLEHAN BOUNDARY LOGIC ---
    for i, line_text in enumerate(all_lines):
        if "TAJUK PEROLEHAN" in line_text.upper():
            collected_title = []
            remainder = re.sub(r'(?i)TAJUK PEROLEHAN\s*:?\s*', '', line_text).strip()
            if remainder: collected_title.append(remainder)
                
            for j in range(i + 1, min(i + 5, len(all_lines))):
                next_line = all_lines[j].strip()
                next_line_upper = next_line.upper()
                if "TARIKH PERMOHONAN" in next_line_upper or "DITERIMA" in next_line_upper: break 
                collected_title.append(next_line)
            
            if collected_title: extracted_data["TAJUK_PEROLEHAN"] = " ".join(collected_title).strip()
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
                
                time.sleep(1) 
            
            if extracted_records:
                df = pd.DataFrame(extracted_records)
                
                if 'NAMA_SYARIKAT' in df.columns:
                    df['NAMA_SYARIKAT'] = df['NAMA_SYARIKAT'].astype(str).str.replace(r'^\d{6}-\d{2}-\d{4}\s*', '', regex=True)

                if 'TARIKH' in df.columns:
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
    
# --- [NANDATANGAN - IMAGE PROCESSING FUNCTION] ---
def process_signature(uploaded_file):
    """
    Takes an uploaded image file (bytes), removes the white background, 
    makes it semi-transparent, and sharpens it.
    """
    try:
        # Open image and convert to RGBA to allow transparency
        img = Image.open(uploaded_file).convert("RGBA")
        # FIXED: getdata() instead of get_data()
        datas = img.getdata()

        newData = []
        # Background Removal Logic (Pillow)
        tolerance = 230
        for item in datas:
            if item[0] > tolerance and item[1] > tolerance and item[2] > tolerance:
                # Replace white pixel with transparent pixel
                newData.append((255, 255, 255, 0))
            else:
                # Translucence: Make non-white pixels semi-transparent
                newData.append((item[0], item[1], item[2], 180)) 
        
        img.putdata(newData)
        
        # Sharpening Filter
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.SHARPEN) # Sharpen twice for boldness

        # Save to a byte stream
        img_bytes = io.BytesIO()
        img.save(img_bytes, format='PNG') # Must be PNG to preserve transparency
        img_bytes.seek(0)
        return img_bytes
    except Exception as e:
        st.error(f"Error processing signature image: {e}")
        return None
# ---------------------------------------------------
                
                # --- ZIP FILE NAMING LOGIC ---
                first_company_full = str(edited_df.iloc[0].get('NAMA_SYARIKAT', 'Syarikat')).strip()
                first_company_word = re.sub(r'[^A-Za-z0-9]', '', first_company_full.split()[0]) if first_company_full else "Syarikat"
                
                current_count = st.session_state.zip_counter
                zip_filename = f"{first_company_word}_{current_count}.zip"
                
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for index, row in edited_df.iterrows():
                        doc = DocxTemplate(template_file)
                        context = row.dropna().to_dict()
                        
                        # --- [NANDATANGAN - ATTACH IMAGE TO CONTEXT] ---
                        # We must reset the stream pointer for EVERY document loop
                        if processed_sign_stream and signature_file:
                            processed_sign_stream.seek(0)
                            # Create the InlineImage object needed by docxtpl
                            # We set a fixed width (e.g., 45mm), which keeps aspect ratio
                            sign_inline = InlineImage(doc, processed_sign_stream, width=Mm(45))
                            context['sign'] = sign_inline
                        # -----------------------------------------------

                        doc.render(context)
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # --- INDIVIDUAL DOCX NAMING LOGIC ---
                        row_company_full = str(row.get('NAMA_SYARIKAT', 'Syarikat')).strip()
                        row_company_word = re.sub(r'[^A-Za-z0-9]', '', row_company_full.split()[0]) if row_company_full else "Syarikat"
                        
                        tajuk_text = str(row.get('TAJUK_PEROLEHAN', ''))
                        all_digits = re.sub(r'\D', '', tajuk_text) 
                        last_5_digits = all_digits[-5:] if all_digits else "00000"
                            
                        final_docx_name = f"{row_company_word}_{last_5_digits}.docx"
                        zip_file.writestr(final_docx_name, doc_io.getvalue())
                
                # Increment our running memory counter for the next batch!
                st.session_state.zip_counter += 1
                
                st.download_button(
                    label=f"⬇️ Download Documents ({zip_filename})",
                    data=zip_buffer.getvalue(),
                    file_name=zip_filename,
                    mime="application/zip"
                )

# --- 6. START OVER BUTTON ---
st.markdown("---") 
col1, col2, col3, col4 = st.columns([1, 1, 1, 1])

with col4: 
    if st.button("🔄 Start Over", type="primary", use_container_width=True):
        saved_counter = st.session_state.zip_counter
        
        st.session_state.clear()
        
        st.session_state.zip_counter = saved_counter
        st.session_state.uploader_key = str(time.time())
        st.rerun()
