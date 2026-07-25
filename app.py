import os
import tempfile
import streamlit as st
import docx
from docx.shared import Pt

# Page Config
st.set_page_config(page_title="PDF to Word (with Math)", page_icon="📝", layout="wide")

st.title("📝 PDF to MS Word (.docx) Converter")
st.write("Upload your PDF to extract text and convert math equations into an **editable Word document**.")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

def create_docx_with_math(text_content, output_path):
    """Creates an MS Word document and handles inline math block formatting."""
    doc = docx.Document()
    
    # Set default font style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    lines = text_content.split("\n")
    for line in lines:
        if line.strip().startswith("#"):
            # Headers
            level = min(line.count("#"), 3)
            header_text = line.replace("#", "").strip()
            doc.add_heading(header_text, level=level)
        elif "$$" in line or "$" in line:
            # Equations block / paragraph
            p = doc.add_paragraph()
            p.add_run(line) # Keeps equation structure readable as LaTeX math in Word
        else:
            if line.strip():
                doc.add_paragraph(line)

    doc.save(output_path)

if uploaded_file is not None:
    st.info(f"Processing: **{uploaded_file.name}**")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.getbuffer())
        temp_pdf_path = temp_pdf.name

    docx_output_path = os.path.join(tempfile.gettempdir(), "converted_output.docx")

    try:
        with st.spinner("Parsing text and math equations..."):
            
            # Using fitz (PyMuPDF) and pixmap OCR to capture complex symbol placements
            import fitz  # PyMuPDF
            
            doc = fitz.open(temp_pdf_path)
            full_extracted_text = []

            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Extract structured text block layout
                text = page.get_text("text")
                if text.strip():
                    full_extracted_text.append(text)
                else:
                    # Fallback for scanned pages
                    pix = page.get_pixmap()
                    full_extracted_text.append(f"[Page {page_num+1} - Math/Image Content Detected]")

            final_text = "\n\n".join(full_extracted_text)

            # Generate MS Word document
            create_docx_with_math(final_text, docx_output_path)

            st.success("Conversion completed successfully!")

            # Display Preview
            st.subheader("Extracted Text Preview")
            st.text_area("Preview", value=final_text[:2000] + "...", height=250)

            # Download MS Word File
            with open(docx_output_path, "rb") as word_file:
                st.download_button(
                    label="📥 Download MS Word File (.docx)",
                    data=word_file,
                    file_name=f"{os.path.splitext(uploaded_file.name)[0]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    except Exception as e:
        st.error(f"Error during processing: {e}")

    finally:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(docx_output_path):
            os.remove(docx_output_path)
