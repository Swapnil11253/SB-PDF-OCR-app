import streamlit as st
import tempfile
import os
import fitz  # PyMuPDF
from docx import Document

# ======================================================================
# CONFIGURATION & CONSTANTS
# ======================================================================
EXTRACTION_MODES = {
    "Standard (Text + Equations)": "standard",
    "OCR Focused (Scanned Documents)": "ocr",
    "Fast Text Only": "fast"
}

LANG_MAP = {
    "English": "eng",
    "Hindi / Devanagari": "hin",
    "Multilingual": "mul"
}

# ======================================================================
# REAL PDF PROCESSING PIPELINE
# ======================================================================
def process_pdf_document(input_path, output_path, mode, lang):
    """
    Reads text and pages from the uploaded PDF and writes them into a .docx file.
    """
    doc = Document()
    doc.add_heading('Converted Document Output', 0)
    
    # Open the uploaded PDF using PyMuPDF (fitz)
    pdf_document = fitz.open(input_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        
        # Add Page Heading
        doc.add_heading(f'Page {page_num + 1}', level=2)
        
        if text.strip():
            doc.add_paragraph(text)
        else:
            doc.add_paragraph("[No readable text found on this page]")
            
    pdf_document.close()
    doc.save(output_path)
    return True

# ======================================================================
# STREAMLIT USER INTERFACE
# ======================================================================
def run_streamlit_app():
    st.set_page_config(page_title="Document Processor & Converter", layout="wide", page_icon="📄")
    
    st.title("📄 PDF / Document Processing Pipeline")
    st.write("Convert complex PDFs with text into editable Word (.docx) documents.")

    # Sidebar Options
    st.sidebar.header("⚙️ Configuration Options")
    
    ocr_mode = st.sidebar.selectbox(
        "Extraction Mode",
        options=list(EXTRACTION_MODES.keys()),
        index=0
    )
    
    selected_lang = st.sidebar.selectbox(
        "Language Priority",
        options=list(LANG_MAP.keys()),
        index=0
    )

    # Main Upload Area
    uploaded_file = st.file_uploader("Upload PDF File", type=["pdf"])

    if uploaded_file is not None:
        st.info(f"📁 File uploaded: **{uploaded_file.name}**")
        
        if st.button("🚀 Start Processing", type="primary"):
            with st.spinner("Processing document... Please wait."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_pdf_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(input_pdf_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    output_docx_path = os.path.join(temp_dir, "Converted_Document.docx")
                    
                    try:
                        # Process PDF
                        process_pdf_document(
                            input_pdf_path, 
                            output_docx_path, 
                            EXTRACTION_MODES[ocr_mode], 
                            LANG_MAP[selected_lang]
                        )
                        
                        st.success("✨ Conversion completed successfully!")
                        
                        # Read converted file for user download
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
