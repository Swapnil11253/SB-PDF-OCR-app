import sys
import os
import re
import glob
import shutil
import subprocess

# ================================
# 1. ENVIRONMENT & DEPENDENCY LOCK
# ================================
print("📦 Checking Dependencies...")
install_cmd = [
    sys.executable, "-m", "pip", "install", "-q",
    "transformers==4.45.2",
    "marker-pdf>=1.0.0",
    "surya-ocr>=0.6.0",
    "click==8.1.8",
    "pypdf",
    "pdf2image"
]

# System Environment Settings
os.environ["INFERENCE_BACKEND"] = "pt"   # Docker Bypass
os.environ["FORCE_OCR"] = "true"         # Force OCR on all pages

subprocess.run("apt-get update -y && apt-get install -y poppler-utils pandoc", shell=True, check=False)

from google.colab import files, userdata

# ================================
# 2. CONFIG & FILE UPLOAD
# ================================
USE_LLM = True
try:
    GEMINI_API_KEY = userdata.get("GEMINI_API_KEY")
except Exception:
    GEMINI_API_KEY = None

if USE_LLM and not GEMINI_API_KEY:
    print("⚠️ GEMINI_API_KEY nahi mila. USE_LLM ko False set kar rahe hain.")
    USE_LLM = False

print("\n🔄 Upload your PDF file:")
uploaded = files.upload()
raw_filename = list(uploaded.keys())[0]

pdf_filename = "doc_input.pdf"
if os.path.exists(pdf_filename):
    os.remove(pdf_filename)
os.rename(raw_filename, pdf_filename)

output_dir = "marker_out"
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir, exist_ok=True)

# ================================
# 3. RUN MARKER PARSER ENGINE
# ================================
print("\n⚡ Running Marker Layout Engine...")

# Cleaned CLI Command Syntax using --output_dir flag
cmd = [
    "marker_single",
    pdf_filename,
    "--output_dir", output_dir,
    "--debug"
]

if USE_LLM and GEMINI_API_KEY:
    os.environ["GEMINI_API_KEY"] = GEMINI_API_KEY
    cmd.append("--use_llm")

result = subprocess.run(cmd, capture_output=True, text=True, env=os.environ)

# Debug Log Save
with open("marker_full_log.txt", "w") as f:
    f.write("=== STDOUT ===\n" + result.stdout + "\n\n=== STDERR ===\n" + result.stderr)

md_files = glob.glob(os.path.join(output_dir, "**", "*.md"), recursive=True)
if not md_files:
    print("❌ Error: Markdown output generate nahi hua. Full log Check Karein:")
    print(result.stderr[-3000:])
    raise SystemExit(1)

md_path = md_files[0]
md_dir = os.path.dirname(md_path)
print(f"🎉 Success! Generated Markdown path: {md_path}")

# ================================
# 4. MCQ FORMATTER & WORD CONVERSION
# ================================
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

with open(md_path, "r", encoding="utf-8") as f:
    formatted_md = split_inline_mcq_options(f.read())

with open(md_path, "w", encoding="utf-8") as f:
    f.write(formatted_md)

# Convert Markdown to editable Word (.docx)
base_name = os.path.splitext(raw_filename)[0]
docx_name = f"{base_name}.docx"

pandoc_cmd = [
    "pandoc",
    os.path.basename(md_path),
    "-f", "markdown+tex_math_dollars+tex_math_single_backslash+raw_tex+pipe_tables+grid_tables+raw_html",
    "-o", docx_name,
]

result = subprocess.run(pandoc_cmd, cwd=md_dir, capture_output=True, text=True)

docx_path = os.path.join(md_dir, docx_name)
if result.returncode == 0 and os.path.exists(docx_path):
    print(f"\n🎉 Word Document Ready & Downloading: {docx_path}")
    files.download(docx_path)
else:
    print("❌ Pandoc conversion failed:")
    print(result.stderr)
