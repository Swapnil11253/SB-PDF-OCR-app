import os
import tempfile
import streamlit as st

# Safe import check for python-docx
try:
    import docx
    from docx.shared import Pt
except ModuleNotFoundError:
    st.error("Missing dependency! Please add 'python-docx' to your requirements.txt file.")
    st.stop()

# Page Config
st.set_page_config(page_title="PDF to Word Converter", page_icon="📝", layout="wide")

st.title("📝 PDF to MS Word (.docx) Converter")
st.write("Upload a PDF to convert its text and math formulas into an **editable Word document**.")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

def create_docx(text_content, output_path):
    """Generates an MS Word document from extracted text."""
    doc = docx.Document()
    
    # Set default document styling
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    # Write text line by line to Word document
    lines = text_content.split("\n")
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
        
        # Format headings or math equations
        if cleaned_line.startswith("#"):
            level = min(cleaned_line.count("#"), 3)
            doc.add_heading(cleaned_line.replace("#", "").strip(), level=level)
        else:
            p = doc.add_paragraph()
            p.add_run(cleaned_line)

    doc.save(output_path)

if uploaded_file is not None:
    st.info(f"File uploaded: **{uploaded_file.name}**")
    
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.getbuffer())
        temp_pdf_path = temp_pdf.name

    docx_output_path = os.path.join(tempfile.gettempdir(), "converted_output.docx")

    try:
        with st.spinner("Extracting text and equations from PDF..."):
            import fitz  # PyMuPDF

            doc = fitz.open(temp_pdf_path)
            extracted_pages = []

            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text("text")
                
                if text.strip():
                    extracted_pages.append(text)

            final_text = "\n\n".join(extracted_pages)

            if not final_text.strip():
                st.warning("No readable text found in PDF. The document might be scanned images.")
            else:
                # Build Word document
                create_docx(final_text, docx_output_path)
                st.success("Conversion successful!")

                # Preview output
                st.subheader("Extracted Text Preview")
                st.text_area("Preview", value=final_text[:1500] + "\n\n...", height=250)

                # Download Button
                with open(docx_output_path, "rb") as word_file:
                    st.download_button(
                        label="📥 Download MS Word File (.docx)",
                        data=word_file,
                        file_name=f"{os.path.splitext(uploaded_file.name)[0]}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

    except Exception as e:
        st.error(f"Error during conversion: {e}")

    finally:
        # Clean up temporary files
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(docx_output_path):
            os.remove(docx_output_path)
