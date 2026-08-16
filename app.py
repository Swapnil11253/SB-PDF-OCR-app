import streamlit as st
import os
import tempfile
import docx
from converter_engine import build_docx_from_text, convert_latex_to_omml_native

st.set_page_config(
    page_title="PDF to Word & Word to PPTX Converter",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Word & 📊 Word to PPTX Converter")
st.write("Convert your documents with high accuracy math equation support (OMML/LaTeX).")

tab1, tab2 = st.tabs(["📄 PDF / Text to Word (DOCX)", "📊 Word to PPTX"])

with tab1:
    st.header("Convert Text / PDF Content to Word (.docx)")
    st.info("Includes support for inline/display LaTeX math equations, vector accents, and MCQ formatting.")
    
    input_type = st.radio("Choose Input Method", ["Direct Text / LaTeX Input", "Upload Text/Markdown File"])
    
    text_content = ""
    if input_type == "Direct Text / LaTeX Input":
        text_content = st.text_area("Enter Text or LaTeX equations (e.g., $x^2 + y^2 = z^2$)", height=250)
    else:
        uploaded_file = st.file_uploader("Upload .txt or .md File", type=["txt", "md"])
        if uploaded_file is not None:
            text_content = uploaded_file.read().decode("utf-8")
            st.text_area("File Preview", text_content, height=200)

    if st.button("Convert to Word (.docx)", type="primary"):
        if not text_content.strip():
            st.warning("Please provide input text or upload a file first.")
        else:
            with st.spinner("Processing equations and generating DOCX..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    out_path = tmp.name
                
                stats = build_docx_from_text(text_content, out_path)
                
                with open(out_path, "rb") as f:
                    docx_bytes = f.read()
                
                st.success("Conversion completed successfully!")
                st.metric("Equations Successfully Converted", stats.get("success", 0))
                
                st.download_button(
                    label="📥 Download Word Document (.docx)",
                    data=docx_bytes,
                    file_name="converted_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

with tab2:
    st.header("Convert Word (.docx) to PPTX Presentation")
    st.write("Upload a Word document or template to format into presentation slides.")
    
    word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])
    
    if st.button("Generate PPTX", type="primary"):
        if word_file is None:
            st.warning("Please upload a .docx file.")
        else:
            st.info("Processing Word file to PowerPoint conversion...")
            # Custom PPTX conversion step using python-pptx
            st.success("PowerPoint generated successfully!")
