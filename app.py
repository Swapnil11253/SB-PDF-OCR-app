"""
PDF -> Word (native OMML equations) converter
-----------------------------------------------
Streamlit port of the original Colab notebook pipeline:

  PDF --(marker_single)--> Markdown + images/tables
      --(pandoc)---------> .docx (native tables, embedded images, most equations)
      --(equation repair)-> .docx (any leftover raw "$...$" text turned into real OMML)

Run locally:
    pip install -r requirements.txt
    # + system packages: pandoc, poppler-utils  (see packages.txt / README)
    streamlit run app.py

NOTE on marker-pdf: it downloads ~2-3 GB of model weights on first run and is
much faster with a GPU. On Streamlit Community Cloud (CPU-only) conversion
will be slow, especially with --force_ocr. Keep that in mind for large PDFs.
"""

import os
import glob
import shutil
import subprocess
import tempfile
import traceback
from pathlib import Path

import streamlit as st

# ----------------------------------------------------------------------------
# Page setup
# ----------------------------------------------------------------------------
st.set_page_config(page_title="PDF → Word (OMML Equations)", page_icon="📄", layout="centered")

st.title("📄 PDF → Word Converter")
st.caption("Native editable equations (OMML) · Native tables · Embedded images")

with st.expander("ℹ️ Ye app kya karta hai (kaise kaam karta hai)"):
    st.markdown(
        """
1. **Marker** aapke PDF ko layout-aware tareeke se Markdown mein convert karta hai
   (text, tables, images, equations sab alag-alag detect karke).
2. **Pandoc** us Markdown ko `.docx` mein convert karta hai — equations native
   Word equations (OMML) ban jaate hain, tables editable Word tables, aur
   diagrams real embedded images.
3. **Equation-repair step** un equations ko dhoondh kar fix karta hai jinhe
   Pandoc parse nahi kar paaya (wo silently raw `$...$` text ke roop mein reh
   jaate hain) — LaTeX → MathML → OMML convert karke unhe bhi real Word
   equation bana deta hai.
        """
    )

# ----------------------------------------------------------------------------
# Cached one-time environment check (does NOT install anything at runtime;
# assumes marker-pdf / pandoc / poppler-utils are already available in the
# environment via requirements.txt + packages.txt — see README below).
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def check_dependencies():
    missing = []
    for exe in ["pandoc", "marker_single", "pdftoppm"]:
        if shutil.which(exe) is None:
            missing.append(exe)
    return missing


try:
    missing_tools = check_dependencies()
except Exception:
    missing_tools = []
    st.warning("⚠️ Environment check nahi ho paaya, aage badh rahe hain — agar conversion fail ho to error details neeche milenge.")

if missing_tools:
    st.error(
        "⚠️ Ye system tools missing hain: **" + ", ".join(missing_tools) + "**.\n\n"
        "Ye app khud inhe runtime par install nahi karta (Colab ke `!pip install` / "
        "`!apt-get install` jaisa yahan nahi chalta). Deploy karte waqt "
        "`requirements.txt` aur `packages.txt` (ya Dockerfile) mein inhe add karein — "
        "neeche README section mein poori list di gayi hai."
    )

st.caption(
    "⚠️ Marker-pdf models load hone mein 1-2+ GB RAM lagti hai. Free/low-RAM hosting "
    "(jaise Streamlit Community Cloud ka free tier) par bade PDF ya paralel conversions "
    "resource limit cross karke app crash kar sakte hain — chhote PDF se test karein."
)

# ----------------------------------------------------------------------------
# Session state defaults
# ----------------------------------------------------------------------------
if "final_docx_path" not in st.session_state:
    st.session_state.final_docx_path = None
if "final_docx_name" not in st.session_state:
    st.session_state.final_docx_name = None
if "log_lines" not in st.session_state:
    st.session_state.log_lines = []


def log(msg: str):
    st.session_state.log_lines.append(msg)


# ----------------------------------------------------------------------------
# Core pipeline (mirrors the notebook, minus Colab-specific upload/download)
# ----------------------------------------------------------------------------
def run_marker(pdf_path: str, output_dir: str, force_ocr: bool, status, page_range: str = None):
    cmd = ["marker_single", pdf_path, "--output_dir", output_dir, "--output_format", "markdown"]
    if force_ocr:
        cmd.append("--force_ocr")
    if page_range:
        cmd.extend(["--page_range", page_range])
    # NOTE: --debug hata diya (default se off) — ye har page ki debug/layout
    # images bhi save karta hai, jo CPU-only / low-RAM hosting (jaise
    # Streamlit Community Cloud free tier) par memory aur disk dono zyada
    # use karta hai aur silent OOM-crash ka ek common reason hai.

    # Low-memory environment: threading libraries (torch/BLAS/OMP) spawn ek
    # thread per CPU core by default, jo RAM-constrained container par
    # peak memory usage badha deta hai. Single-threaded chalane se peak
    # memory kam predictable/lower rehti hai (thoda slow ho sakta hai).
    env = os.environ.copy()
    env.setdefault("OMP_NUM_THREADS", "1")
    env.setdefault("MKL_NUM_THREADS", "1")
    env.setdefault("TORCH_NUM_THREADS", "1")

    status.write(f"🔧 Running: `{' '.join(cmd)}`")
    result = subprocess.run(cmd, capture_output=True, text=True, env=env)

    log_path = os.path.join(output_dir, "marker_full_log.txt")
    with open(log_path, "w") as f:
        f.write("=== STDOUT ===\n" + result.stdout + "\n\n=== STDERR ===\n" + result.stderr)

    if result.returncode != 0:
        raise RuntimeError(f"Marker execution failed:\n{result.stderr[-3000:]}")

    md_files = glob.glob(os.path.join(output_dir, "**", "*.md"), recursive=True)
    if not md_files:
        raise RuntimeError(f"No .md file found after marker run.\n{result.stderr[-3000:]}")

    return md_files[0], log_path


def run_pandoc(md_path: str, docx_name: str, status):
    md_dir = os.path.dirname(md_path)
    pandoc_cmd = [
        "pandoc",
        os.path.basename(md_path),
        "-f",
        "markdown+tex_math_dollars+tex_math_single_backslash+raw_tex+pipe_tables+grid_tables+raw_html",
        "-o",
        docx_name,
    ]
    status.write(f"🔧 Running (in `{md_dir}`): `{' '.join(pandoc_cmd)}`")
    result = subprocess.run(pandoc_cmd, cwd=md_dir, capture_output=True, text=True)

    docx_path = os.path.join(md_dir, docx_name)
    if result.returncode != 0 or not os.path.exists(docx_path):
        raise RuntimeError(f"Pandoc conversion failed.\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}")

    return docx_path, result.stderr


# ---- Equation repair (LaTeX -> MathML -> OMML) ----
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _latex_to_omath_element(latex_src):
    import latex2mathml.converter as _l2m
    import mathml2omml as _mathml2omml
    from lxml import etree as _etree

    mathml = _l2m.convert(latex_src.strip())
    omml_str = _mathml2omml.convert(mathml)  # '<m:oMath>...</m:oMath>', no ns decl
    wrapped = f'<m:root xmlns:m="{_M_NS}">{omml_str}</m:root>'
    return _etree.fromstring(wrapped.encode("utf-8"))[0]


def repair_docx_equations(docx_in: str, docx_out: str):
    """
    Scan docx_in for leftover raw '$...$' / '$$...$$' text runs (pandoc's
    parse-failure fallback) and replace each with a real OMML equation.
    Writes the fixed file to docx_out. Returns (fixed_count, failed_count, notes).
    """
    import re as _re
    import zipfile as _zipfile
    from lxml import etree as _etree

    notes = []
    tmp_dir = docx_out + "_tmpext"
    shutil.rmtree(tmp_dir, ignore_errors=True)
    with _zipfile.ZipFile(docx_in) as z:
        z.extractall(tmp_dir)

    doc_path = f"{tmp_dir}/word/document.xml"
    tree = _etree.parse(doc_path)
    root = tree.getroot()

    fixed, failed = 0, 0
    for t_el in root.iter(f"{{{_W_NS}}}t"):
        text = (t_el.text or "").strip()
        m = _re.fullmatch(r"\$\$(.+?)\$\$", text, _re.DOTALL) or _re.fullmatch(r"\$(.+?)\$", text, _re.DOTALL)
        if not m:
            continue
        latex_src = next(g for g in m.groups() if g is not None)
        try:
            omath_el = _latex_to_omath_element(latex_src)
        except Exception as e:
            failed += 1
            notes.append(f"[SKIP] could not repair: {latex_src[:70]!r} -> {e}")
            continue
        run_el = t_el.getparent()
        run_el.getparent().replace(run_el, omath_el)
        fixed += 1
        notes.append(f"[FIXED] {latex_src[:70]!r}")

    tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)

    with _zipfile.ZipFile(docx_out, "w", _zipfile.ZIP_DEFLATED) as zout:
        for base, _, files_ in os.walk(tmp_dir):
            for fn in files_:
                full = os.path.join(base, fn)
                zout.write(full, os.path.relpath(full, tmp_dir))
    shutil.rmtree(tmp_dir, ignore_errors=True)

    return fixed, failed, notes


# ----------------------------------------------------------------------------
# UI — upload + options
# ----------------------------------------------------------------------------
uploaded_pdf = st.file_uploader("Apna PDF upload karein", type=["pdf"])

total_pages = None
if uploaded_pdf is not None:
    try:
        from pypdf import PdfReader

        total_pages = len(PdfReader(uploaded_pdf).pages)
        uploaded_pdf.seek(0)  # reader ne stream consume kar liya, wapas rewind
    except Exception:
        total_pages = None

col1, col2 = st.columns(2)
with col1:
    force_ocr = st.checkbox(
        "Force OCR",
        value=True,
        help="PDF ka embedded text layer ignore karke poora document fresh OCR se padhega. "
        "Missing/incomplete paragraphs fix karta hai, lekin slower hai (aur zyada RAM leta hai).",
    )
with col2:
    do_repair = st.checkbox(
        "Equation-repair step chalayein",
        value=True,
        help="Jo equations Pandoc parse nahi kar paya, unhe LaTeX → MathML → OMML "
        "convert karke real Word equation banata hai.",
    )

page_range = None
if total_pages:
    st.caption(f"📄 Is PDF mein **{total_pages} pages** hain.")
    if total_pages > 8:
        st.warning(
            "⚠️ Ye PDF kaafi bada hai. Marker ke OCR/layout models chalane ke liye "
            "kaafi RAM chahiye hoti hai — free/low-RAM hosting par bade PDFs beech mein "
            "hi crash ('Error running app') kar sakte hain. Neeche page-range choose "
            "karke chhote batch mein convert karna safer rahega."
        )
    limit_pages = st.checkbox("Sirf kuch pages convert karein (RAM bachane ke liye)", value=total_pages > 8)
    if limit_pages:
        start_p, end_p = st.slider(
            "Page range (0-indexed, dono taraf inclusive)",
            min_value=0,
            max_value=max(total_pages - 1, 0),
            value=(0, min(total_pages - 1, 7)),
        )
        page_range = f"{start_p}-{end_p}"
        st.caption(f"👉 Convert hoga: page {start_p} se {end_p} (kul {end_p - start_p + 1} pages)")

convert_clicked = st.button("🚀 Convert to Word", type="primary", disabled=uploaded_pdf is None or bool(missing_tools))

# ----------------------------------------------------------------------------
# Run pipeline
# ----------------------------------------------------------------------------
if convert_clicked and uploaded_pdf is not None:
    st.session_state.final_docx_path = None
    st.session_state.final_docx_name = None
    st.session_state.log_lines = []

    work_dir = tempfile.mkdtemp(prefix="pdf2word_")
    pdf_path = os.path.join(work_dir, "doc_input.pdf")
    with open(pdf_path, "wb") as f:
        f.write(uploaded_pdf.getbuffer())

    output_dir = os.path.join(work_dir, "marker_out")
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(uploaded_pdf.name)[0]
    docx_name = f"{base_name}.docx"

    try:
        with st.status("⚡ Step 1/3 — Marker: PDF → Markdown...", expanded=True) as status:
            md_path, marker_log_path = run_marker(pdf_path, output_dir, force_ocr, status, page_range=page_range)
            status.write(f"💾 Markdown ready: `{os.path.basename(md_path)}`")

            # quick word-count sanity check, same as the notebook
            try:
                from pypdf import PdfReader

                reader = PdfReader(pdf_path)
                raw_text = "\n".join((page.extract_text() or "") for page in reader.pages)
                raw_word_count = len(raw_text.split())
                with open(md_path, "r", encoding="utf-8") as f:
                    md_word_count = len(f.read().split())
                status.write(f"📊 Word count — Original PDF: {raw_word_count} | Marker output: {md_word_count}")
                if raw_word_count > 0 and md_word_count < 0.7 * raw_word_count:
                    status.write("⚠️ Marker output PDF ke raw text se kaafi kam hai — kuch text drop ho sakta hai. Full log neeche download karein.")
            except Exception as e:
                status.write(f"(Word-count diagnostic skip: {e})")

            status.update(label="✅ Step 1/3 done — Markdown generated", state="running")

        with st.status("📝 Step 2/3 — Pandoc: Markdown → .docx...", expanded=True) as status:
            docx_path, pandoc_stderr = run_pandoc(md_path, docx_name, status)
            if pandoc_stderr.strip():
                status.write("⚠️ Pandoc warnings (ye equations aage repair step mein fix ho sakte hain):")
                status.write(f"```\n{pandoc_stderr[:2000]}\n```")
            status.update(label="✅ Step 2/3 done — .docx created", state="running")

        final_path = docx_path
        if do_repair:
            with st.status("🔧 Step 3/3 — Repairing unparsed equations...", expanded=True) as status:
                repaired_path = docx_path.replace(".docx", "_repaired.docx")
                fixed_n, failed_n, notes = repair_docx_equations(docx_path, repaired_path)
                for n in notes[:50]:
                    status.write(n)
                status.write(f"**Equation repair: {fixed_n} fixed, {failed_n} still unrepaired.**")
                if failed_n:
                    status.write("⚠️ Kuch equations abhi bhi auto-fix nahi ho paaye — doc kholke literal '$' search karke manually check karein.")
                final_path = repaired_path if os.path.exists(repaired_path) else docx_path
                status.update(label="✅ Step 3/3 done", state="complete")
        else:
            st.info("Equation-repair step skip kiya gaya (checkbox off tha).")

        st.session_state.final_docx_path = final_path
        st.session_state.final_docx_name = os.path.basename(final_path)
        st.session_state.marker_log_path = marker_log_path
        st.success("🎉 Conversion complete!")

    except Exception as e:
        st.error(f"❌ Conversion failed: {e}")
        with st.expander("🔍 Full error details (technical)"):
            st.code(traceback.format_exc())
        st.info(
            "Agar ye baar-baar ho raha hai aur koi Python error yahan nahi dikh raha "
            "(sirf blank/generic 'Error running app' page dikhta hai), to zyada chance "
            "hai ki hosting environment ki RAM/disk limit cross ho rahi hai — marker-pdf "
            "ke models kaafi memory-heavy hain. Chhota PDF try karein, ya zyada RAM wale "
            "plan / apne server par deploy karein."
        )

# ----------------------------------------------------------------------------
# Download buttons
# ----------------------------------------------------------------------------
if st.session_state.final_docx_path and os.path.exists(st.session_state.final_docx_path):
    with open(st.session_state.final_docx_path, "rb") as f:
        st.download_button(
            "📥 Download Word document (.docx)",
            data=f.read(),
            file_name=st.session_state.final_docx_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    marker_log_path = st.session_state.get("marker_log_path")
    if marker_log_path and os.path.exists(marker_log_path):
        with open(marker_log_path, "rb") as f:
            st.download_button(
                "📄 Download full Marker log (debugging ke liye)",
                data=f.read(),
                file_name="marker_full_log.txt",
                mime="text/plain",
            )
