import streamlit as st
import tempfile
import os

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
# CORE PROCESSING PIPELINE (PLACEHOLDER FOR YOUR ENGINE LOGIC)
# ======================================================================
def process_pdf_document(input_path, output_path, mode, lang):
    """
    Placeholder function for your PDF processing engine logic.
    Replace or connect your existing core functions (e.g., PyMuPDF, docx builder) here.
    """
    from docx import Document
    doc = Document()
    doc.add_heading('Converted Document Output', 0)
    doc.add_paragraph(f"Processed file: {os.path.basename(input_path)}")
    doc.add_paragraph(f"Extraction Mode: {mode}")
    doc.add_paragraph(f"Language Setting: {lang}")
    doc.save(output_path)
    return True

# ======================================================================
# STREAMLIT USER INTERFACE
# ======================================================================
def run_streamlit_app():
    st.set_page_config(page_title="Document Processor & Converter", layout="wide", page_icon="📄")
    
    st.title("📄 PDF / Document Processing Pipeline")
    st.write("Convert complex PDFs with LaTeX equations into editable Word (.docx) and PowerPoint (.pptx) documents.")

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
                # Save uploaded file to temporary directory
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_pdf_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(input_pdf_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    output_docx_path = os.path.join(temp_dir, "Converted_Document.docx")
                    
                    try:
                        # Process document
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
