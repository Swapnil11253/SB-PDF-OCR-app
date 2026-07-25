import os
import re
import glob
import shutil
import subprocess
import streamlit as st

# Streamlit Page Config
st.set_page_config(page_title="PDF to Word (OCR) Converter", page_icon="📄", layout="centered")

st.title("📄 PDF to Word (OCR) Converter")
st.write("PDF upload karein, Marker engine se extract karein aur formatted Word (.docx) download karein.")

# Set System Environment Variables
os.environ["INFERENCE_BACKEND"] = "pt"
os.environ["FORCE_OCR"] = "true"

# ================================
# 1. CONFIG & API KEY SETTINGS
# ================================
st.sidebar.header("⚙️ Configuration")
use_llm = st.sidebar.checkbox("Use Gemini LLM for OCR Improvement", value=True)

gemini_api_key = None
if use_llm:
    # Check Streamlit secrets first, else fallback to sidebar input
    if "GEMINI_API_KEY" in st.secrets:
        gemini_api_key = st.secrets["GEMINI_API_KEY"]
    else:
        gemini_api_key = st.sidebar.text_input("Enter Gemini API Key:", type="password")
    
    if not gemini_api_key:
        st.sidebar.warning("⚠️ API Key nahi mili. LLM mode off rahega.")
        use_llm = False

# ================================
# 2. FILE UPLOAD
# ================================
uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

# Helper function for MCQ formatting
OPTION_MARKER_RE = re.compile(r'\(\s*([a-dA-D])\s*\)')

def split_inline_mcq_options(md_text: str) -> str:
    out_lines = []
    for line in md_text.split("\n"):
        stripped = line.strip()

        if stripped.startswith("#") or stripped.startswith("!["):
            out_lines.append(line)
            continue

        matches = list(OPTION_MARKER_RE.finditer(line))
        if len(matches) < 1:
            out_lines.append(line)
            continue

        safe_to_split = all(line[:m.start()].count("$") % 2 == 0 for m in matches)
        if not safe_to_split:
            out_lines.append(line)
            continue

        segments = []
        prev_end = 0
        for m in matches:
            segment = line[prev_end:m.start()].strip()
            if segment:
                segments.append(segment)
            prev_end = m.start()
        last_segment = line[prev_end:].strip()
        if last_segment:
            segments.append(last_segment)

        if len(segments) <= 1:
            out_lines.append(line)
        else:
            out_lines.extend(segments)
            out_lines.append("")

    return "\n".join(out_lines)

# Process Button
if uploaded_file is not None:
    if st.button("🚀 Process & Convert PDF"):
        with st.status("Processing PDF...", expanded=True) as status:
            
            # Save uploaded file locally
            pdf_filename = "doc_input.pdf"
            with open(pdf_filename, "wb") as f:
                f.write(uploaded_file.getbuffer())

            output_dir = "marker_out"
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            os.makedirs(output_dir, exist_ok=True)

            # ================================
            # 3. RUN MARKER PARSER ENGINE
            # ================================
            st.write("⚡ Running Marker Layout Engine (isne thoda samay lag sakta hai)...")

            cmd = [
                "marker_single",
                pdf_filename,
                "--output_dir", output_dir,
                "--debug"
            ]

            env_vars = os.environ.copy()
            if use_llm and gemini_api_key:
                env_vars["GEMINI_API_KEY"] = gemini_api_key
                cmd.append("--use_llm")

            result = subprocess.run(cmd, capture_output=True, text=True, env=env_vars)

            # Debug log
            with open("marker_full_log.txt", "w", encoding="utf-8") as f:
                f.write("=== STDOUT ===\n" + result.stdout + "\n\n=== STDERR ===\n" + result.stderr)

            md_files = glob.glob(os.path.join(output_dir, "**", "*.md"), recursive=True)
            if not md_files:
                status.update(label="❌ Error in Marker OCR Processing", state="error")
                st.error("Markdown output generate nahi hua. Error Details:")
                st.code(result.stderr[-2000:])
                st.stop()

            md_path = md_files[0]
            md_dir = os.path.dirname(md_path)

            # ================================
            # 4. MCQ FORMATTER & WORD CONVERSION
            # ================================
            st.write("📝 Formatting MCQs and Converting to Word (.docx)...")

            # Format MCQ lines
            with open(md_path, "r", encoding="utf-8") as f:
                formatted_md = split_inline_mcq_options(f.read())

            with open(md_path, "w", encoding="utf-8") as f:
                f.write(formatted_md)

            # Convert to Docx via Pandoc
            raw_filename = uploaded_file.name
            base_name = os.path.splitext(raw_filename)[0]
            docx_name = f"{base_name}.docx"

            pandoc_cmd = [
                "pandoc",
                os.path.basename(md_path),
                "-f", "markdown+tex_math_dollars+tex_math_single_backslash+raw_tex+pipe_tables+grid_tables+raw_html",
                "-o", docx_name,
            ]

            pandoc_result = subprocess.run(pandoc_cmd, cwd=md_dir, capture_output=True, text=True)
            docx_path = os.path.join(md_dir, docx_name)

            if pandoc_result.returncode == 0 and os.path.exists(docx_path):
                status.update(label="🎉 Conversion Completed Successfully!", state="complete")
                
                # Download Button
                with open(docx_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Word Document (.docx)",
                        data=file,
                        file_name=docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                status.update(label="❌ Pandoc Conversion Failed", state="error")
                st.error("Word conversion fail ho gaya:")
                st.code(pandoc_result.stderr)
