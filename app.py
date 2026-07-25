import os
import tempfile
import streamlit as st

# Set environment variables for OCR/Inference backends directly in Python
os.environ["INFERENCE_BACKEND"] = "pt"   # Docker / PyTorch bypass
os.environ["FORCE_OCR"] = "true"         # Force OCR on all pages

# App UI Configuration
st.set_page_config(page_title="PDF OCR App", page_icon="📄", layout="wide")

st.title("📄 PDF OCR Converter")
st.write("Upload a PDF document to run OCR and extract markdown text.")

# Native Streamlit File Uploader (Replaces google.colab.files)
uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    st.info(f"Processing file: **{uploaded_file.name}**")
    
    # Save the uploaded Streamlit file buffer to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.getbuffer())
        temp_pdf_path = temp_pdf.name

    try:
        with st.spinner("Extracting text and running OCR..."):
            # Import your heavy OCR modules inside the run logic to keep initial load snappy
            import pypdf
            from pdf2image import convert_from_path

            # EXAMPLE: Basic PyPDF extraction pass
            reader = pypdf.PdfReader(temp_pdf_path)
            total_pages = len(reader.pages)
            
            st.success(f"Successfully loaded PDF with {total_pages} page(s).")
            
            extracted_text = ""
            for i, page in enumerate(reader.pages):
                text = page.extract_text() or ""
                extracted_text += f"\n--- Page {i + 1} ---\n" + text

            # If you are using marker-pdf or surya-ocr, invoke them directly here:
            # from marker.convert import convert_single_pdf
            # full_text, images, out_meta = convert_single_pdf(temp_pdf_path, model_lst)

            # Display Output
            st.subheader("Extracted Content")
            st.text_area("Markdown / Extracted Text", value=extracted_text, height=400)

            # Download Button
            st.download_button(
                label="📥 Download Output Text",
                data=extracted_text,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_ocr.txt",
                mime="text/plain"
            )

    except Exception as e:
        st.error(f"An error occurred during processing: {e}")

    finally:
        # Clean up temporary local file
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
