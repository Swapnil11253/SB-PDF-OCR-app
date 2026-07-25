import os
import tempfile
import subprocess
import streamlit as st
import pypandoc

# Page setup
st.set_page_config(page_title="High-Precision Math PDF to Word", page_icon="📐", layout="wide")

st.title("📐 Math PDF to MS Word (.docx) Converter")
st.write("Extracts text and complex **Math Equations (OMML/LaTeX)** using Marker-PDF and Pandoc.")

uploaded_file = st.file_uploader("Upload PDF File", type=["pdf"])

if uploaded_file is not None:
    st.info(f"Processing: **{uploaded_file.name}**")
    
    # Save uploaded PDF to temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.getbuffer())
        temp_pdf_path = temp_pdf.name

    docx_output_path = os.path.join(tempfile.gettempdir(), "converted_math_output.docx")
    md_output_path = os.path.join(tempfile.gettempdir(), "extracted_content.md")

    try:
        with st.spinner("Step 1/2: Running Marker AI to parse equations into LaTeX..."):
            from marker.convert import convert_single_pdf
            from marker.models import load_all_models

            # Load Marker PDF AI Models
            @st.cache_resource
            def load_marker_models():
                return load_all_models()

            models = load_marker_models()
            
            # Extract text + LaTeX math formulas
            full_text, images, out_meta = convert_single_pdf(temp_pdf_path, models)

            # Save extracted markdown text with LaTeX math
            with open(md_output_path, "w", encoding="utf-8") as f:
                f.write(full_text)

        with st.spinner("Step 2/2: Converting LaTeX equations to native Word (OMML) using Pandoc..."):
            # Ensure Pandoc is available
            try:
                pypandoc.convert_file(
                    md_output_path,
                    'docx',
                    outputfile=docx_output_path,
                    extra_args=['--from=markdown+tex_math_dollars+raw_tex']
                )
            except Exception as p_err:
                # Fallback to direct subprocess if pypandoc wrapper fails
                cmd = f"pandoc '{md_output_path}' -o '{docx_output_path}' --from=markdown+tex_math_dollars+raw_tex"
                subprocess.run(cmd, shell=True, check=True)

        st.success("Successfully converted PDF to Word with native editable equations!")

        # Tabs for Preview and Download
        tab1, tab2 = st.tabs(["📄 Math Preview (Rendered)", "📥 Download Word File"])

        with tab1:
            st.markdown(full_text)

        with tab2:
            with open(docx_output_path, "rb") as word_file:
                st.download_button(
                    label="📥 Download Editable MS Word File (.docx)",
                    data=word_file,
                    file_name=f"{os.path.splitext(uploaded_file.name)[0]}_math.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    except Exception as e:
        st.error(f"Error occurred during conversion: {e}")

    finally:
        # Clean up temporary files
        for path in [temp_pdf_path, docx_output_path, md_output_path]:
            if os.path.exists(path):
                os.remove(path)
