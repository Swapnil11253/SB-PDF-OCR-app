import os
import tempfile
import streamlit as st

# Force Marker/Surya settings to preserve equations accurately
os.environ["INFERENCE_BACKEND"] = "pt"
os.environ["FORCE_OCR"] = "true"

st.set_page_config(page_title="PDF Math & OCR Converter", page_icon="📐", layout="wide")

st.title("📐 PDF Converter with Math (OMML / LaTeX) Support")
st.write("Upload your PDF to convert formatted text and mathematical equations into accurate Markdown and LaTeX.")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    st.info(f"Processing file: **{uploaded_file.name}**")
    
    # Save uploaded file to temp path
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.getbuffer())
        temp_pdf_path = temp_pdf.name

    try:
        with st.spinner("Extracting text and converting equations to LaTeX..."):
            
            # Use marker-pdf for high-fidelity document and math equation parsing
            from marker.convert import convert_single_pdf
            from marker.models import load_all_models

            # Load models (cached for session speed)
            @st.cache_resource
            def get_models():
                return load_all_models()

            model_lst = get_models()

            # Execute extraction specifically targeted at capturing equations correctly
            full_text, images, out_meta = convert_single_pdf(temp_pdf_path, model_lst)

            st.success("Extraction complete!")

            # Display Output tabbed between rendered math preview and raw code
            tab1, tab2 = st.tabs(["Formatted Preview (Rendered Math)", "Raw Markdown / LaTeX"])

            with tab1:
                st.markdown(full_text)

            with tab2:
                st.text_area("LaTeX & Markdown Output", value=full_text, height=450)

            # Download Button
            st.download_button(
                label="📥 Download Markdown (.md)",
                data=full_text,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_converted.md",
                mime="text/markdown"
            )

    except Exception as e:
        # Fallback if marker fails or memory limits are hit on free hosting
        st.warning("Advanced Math parser failed. Running lightweight fallback...")
        
        import pypdf
        reader = pypdf.PdfReader(temp_pdf_path)
        extracted_text = ""
        for i, page in enumerate(reader.pages):
            extracted_text += f"\n--- Page {i + 1} ---\n" + (page.extract_text() or "")
            
        st.text_area("Fallback Output (Basic Text)", value=extracted_text, height=400)
        st.error(f"Error Details: {e}")

    finally:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
