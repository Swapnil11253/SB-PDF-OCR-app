import streamlit as st
import tempfile
import os
import fitz  # PyMuPDF
from docx import Document
from PIL import Image
import pytesseract

# Note: For local Windows setup, specify tesseract binary path if needed:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ======================================================================
# CONFIGURATION & CONSTANTS
# ======================================================================
EXTRACTION_MODES = {
    "Standard (Text + OCR)": "standard",
    "OCR Focused (Scanned PDF)": "ocr",
    "Fast Text Only": "fast"
}

LANG_MAP = {
    "English": "eng",
    "Hindi / Devanagari": "hin",
    "English + Hindi": "eng+hin"
}

# ======================================================================
# CORE PROCESSING PIPELINE WITH OCR
# ======================================================================
def process_pdf_document(input_path, output_path, mode, lang):
    doc = Document()
    doc.add_heading('Converted Document Output', 0)
    
    pdf_document = fitz.open(input_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        
        doc.add_heading(f'Page {page_num + 1}', level=2)
        
        # If text is available (Digital PDF)
        if text.strip() and mode != "ocr":
            doc.add_paragraph(text)
        else:
            # Scanned PDF: Extract image and apply OCR
            pix = page.get_pixmap(dpi=300)
            img_path = f"temp_page_{page_num}.png"
            pix.save(img_path)
            
            try:
                img = Image.open(img_path)
                ocr_text = pytesseract.image_to_string(img, lang=lang)
                
                if ocr_text.strip():
                    doc.add_paragraph(ocr_text)
                else:
                    doc.add_paragraph("[OCR could not detect any text on this page]")
            except Exception as ocr_err:
                doc.add_paragraph(f"[OCR Processing Error: Ensure Tesseract is installed. Details: {str(ocr_err)}]")
            finally:
                if os.path.exists(img_path):
                    os.remove(img_path)
            
    pdf_document.close()
    doc.save(output_path)
    return True

# ======================================================================
# STREAMLIT USER INTERFACE
# ======================================================================
def run_streamlit_app():
    st.set_page_config(page_title="Scanned PDF OCR Converter", layout="wide", page_icon="📄")
    
    st.title("📄 Scanned PDF / Document OCR Converter")
    st.write("Extract text from scanned PDF images using OCR and download editable Word (.docx) files.")

    # Sidebar Options
    st.sidebar.header("⚙️ OCR Configuration")
    
    ocr_mode = st.sidebar.selectbox(
        "Extraction Mode",
        options=list(EXTRACTION_MODES.keys()),
        index=1
    )
    
    selected_lang = st.sidebar.selectbox(
        "Language Priority",
        options=list(LANG_MAP.keys()),
        index=0
    )

    # Main Upload Area
    uploaded_file = st.file_uploader("Upload Scanned PDF File", type=["pdf"])

    if uploaded_file is not None:
        st.info(f"📁 File uploaded: **{uploaded_file.name}**")
        
        if st.button("🚀 Start OCR Processing", type="primary"):
            with st.spinner("Extracting text via OCR... Please wait."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_pdf_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(input_pdf_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    output_docx_path = os.path.join(temp_dir, "Converted_Document.docx")
                    
                    try:
                        process_pdf_document(
                            input_pdf_path, 
                            output_docx_path, 
                            EXTRACTION_MODES[ocr_mode], 
                            LANG_MAP[selected_lang]
                        )
                        
                        st.success("✨ OCR Conversion completed successfully!")
                        
                        with open(output_docx_path, "rb") as doc_file:
                            st.download_button(
                                label="📥 Download DOCX Result",
                                data=doc_file,
                                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_converted.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    except Exception as e:
                        st.error(f"❌ Processing failed: {str(e)}")

if __name__ == "__main__":
    run_streamlit_app()
