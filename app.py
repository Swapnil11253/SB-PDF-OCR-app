"""
PDF -> Word Universal Pipeline (Streamlit edition)
====================================================
MinerU (layout/text/table/equation) + pix2tex (equation second-opinion)
+ PaddleOCR-hi (Hindi second-opinion) + pandoc (md -> docx) + OMML repair.

Same 3-tool design as the original Colab notebook, ported to a Streamlit
UI: file_uploader instead of google.colab upload, st.progress/st.status
instead of print(), session_state to persist heavy models + intermediate
artifacts across Streamlit reruns, and st.download_button instead of
files.download().

Run with:  streamlit run app.py
"""

import os
import re
import gc
import glob
import json
import shutil
import zipfile
import subprocess
import tempfile

import numpy as np
import streamlit as st

# ----------------------------------------------------------------------
# System dependency checks — mineru/pandoc/poppler are external binaries,
# not Python packages. Even if they're listed in requirements.txt,
# `mineru` specifically needs its console_script on PATH, and pandoc /
# poppler-utils are apt packages (see packages.txt), not pip packages.
# We check up front instead of letting subprocess.run() crash with a raw
# FileNotFoundError deep in the pipeline.
# ----------------------------------------------------------------------
def check_system_deps():
    checks = {
        "mineru": shutil.which("mineru"),
        "pandoc": shutil.which("pandoc"),
        "pdftoppm (poppler-utils)": shutil.which("pdftoppm"),
    }
    missing = [name for name, path in checks.items() if not path]
    return missing

# ----------------------------------------------------------------------
# Page config
# ----------------------------------------------------------------------
st.set_page_config(page_title="PDF to Word — MinerU + pix2tex + PaddleOCR", layout="wide")

st.title("📄 PDF → Word Universal Pipeline")
st.caption("MinerU (layout/OCR) + pix2tex (equation double-check) + PaddleOCR-hi (Hindi double-check) + pandoc (OMML equations)")

_DEVANAGARI_RE = re.compile(r"[\u0900-\u097F]")
_OPTION_MARKER_RE = re.compile(r"\([A-Ea-e]\)")
_QUESTION_NUM_RE = re.compile(r"^(\d{1,3})\.\s+(.*)$")
_PLAIN_TEXT_RE = re.compile(r"^[A-Za-z0-9\s\.,()]+$")

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


# ----------------------------------------------------------------------
# Cached model loaders — heavy, so load once per Streamlit server process
# ----------------------------------------------------------------------
@st.cache_resource(show_spinner="pix2tex (LatexOCR) model load ho raha hai (sirf ek baar)...")
def load_pix2tex():
    from pix2tex.cli import LatexOCR
    return LatexOCR()


@st.cache_resource(show_spinner="PaddleOCR (Hindi) model load ho raha hai (sirf ek baar)...")
def load_paddle_hi():
    from paddleocr import PaddleOCR
    return PaddleOCR(lang="hi", use_angle_cls=True)


# ----------------------------------------------------------------------
# Pipeline steps (each mirrors a notebook cell)
# ----------------------------------------------------------------------
def pdf_has_significant_hindi(path, threshold_ratio=0.15):
    from pypdf import PdfReader
    reader = PdfReader(path)
    total_chars, devanagari_chars = 0, 0
    for page in reader.pages:
        text = page.extract_text() or ""
        total_chars += len(text)
        devanagari_chars += len(_DEVANAGARI_RE.findall(text))
    if total_chars == 0:
        return False, 0.0
    ratio = devanagari_chars / total_chars
    return ratio >= threshold_ratio, ratio


def run_mineru(pdf_path, out_dir, use_devanagari_mode):
    mineru_bin = shutil.which("mineru")
    if not mineru_bin:
        raise FileNotFoundError(
            "'mineru' command nahi mila is server par. Yeh 'mineru[all]' pip "
            "package install hone ke baad bhi ho sakta hai agar: (1) install "
            "fail/skip ho gaya (build logs check karo), (2) console_script "
            "PATH par nahi hai, ya (3) yeh app aisi hosting (jaise Streamlit "
            "Community Cloud) par chal raha hai jahan itni heavy ML pipeline "
            "(GBs ke models, GPU-preferred) chalti hi nahi. README.md ka "
            "'Setup' section dekho — is app ko apne GPU machine/server par "
            "local chalana recommended hai."
        )
    lang_args = ["-l", "devanagari"] if use_devanagari_mode else []
    cmd = [mineru_bin, "-p", pdf_path, "-o", out_dir, "-b", "pipeline", "-f", "true"] + lang_args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return cmd, result


def locate_mineru_outputs(out_dir):
    md_candidates = glob.glob(os.path.join(out_dir, "**", "*.md"), recursive=True)
    middle_json_candidates = glob.glob(os.path.join(out_dir, "**", "*middle.json"), recursive=True)
    return md_candidates, middle_json_candidates


def collect_spans(middle_json_data):
    """middle.json -> (equation_spans, text_spans, page_dims). Defensive
    walk since exact block nesting can vary slightly by MinerU version."""
    equation_spans = []  # (page_idx, bbox, span_type)
    text_spans = []      # (page_idx, bbox, text)
    page_dims = {}

    def walk_blocks(blocks, page_idx):
        for block in blocks or []:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    span_type = span.get("type", "")
                    bbox = span.get("bbox")
                    if not bbox:
                        continue
                    if span_type in ("inline_equation", "interline_equation"):
                        equation_spans.append((page_idx, bbox, span_type))
                    elif span_type == "text":
                        text_spans.append((page_idx, bbox, span.get("content", "")))
            if block.get("blocks"):
                walk_blocks(block.get("blocks"), page_idx)

    pages = middle_json_data.get("pdf_info", [])
    for page_idx, page in enumerate(pages):
        page_size = page.get("page_size") or page.get("page_size_ori")
        if page_size:
            page_dims[page_idx] = tuple(page_size)
        walk_blocks(page.get("preproc_blocks") or page.get("para_blocks") or page.get("blocks"), page_idx)

    return equation_spans, text_spans, page_dims


class PageRenderer:
    """Wraps pypdfium2 page rendering + bbox crop, same as notebook Cell 6."""

    def __init__(self, pdf_path, page_dims, render_dpi=300):
        import pypdfium2 as pdfium
        self.pdf_doc = pdfium.PdfDocument(pdf_path)
        self.page_dims = page_dims
        self.render_dpi = render_dpi
        self._cache = {}

    def get_page_image(self, page_idx):
        if page_idx not in self._cache:
            page = self.pdf_doc[page_idx]
            bitmap = page.render(scale=self.render_dpi / 72)
            self._cache[page_idx] = bitmap.to_pil()
        return self._cache[page_idx]

    def crop_span(self, page_idx, bbox, pad=4):
        img = self.get_page_image(page_idx)
        ref_dims = self.page_dims.get(page_idx)
        if ref_dims:
            scale_x = img.width / ref_dims[0]
            scale_y = img.height / ref_dims[1]
        else:
            scale_x = scale_y = 1
        x0, y0, x1, y1 = bbox
        left = max(0, int(x0 * scale_x) - pad)
        top = max(0, int(y0 * scale_y) - pad)
        right = min(img.width, int(x1 * scale_x) + pad)
        bottom = min(img.height, int(y1 * scale_y) + pad)
        if right <= left or bottom <= top:
            return None
        return img.crop((left, top, right, bottom))


def run_pix2tex_pass(md_text, equation_spans, renderer, latex_ocr, progress_cb=None):
    import torch

    pix2tex_results = [None] * len(equation_spans)
    total = len(equation_spans)

    for i, (page_idx, bbox, span_type) in enumerate(equation_spans):
        try:
            crop = renderer.crop_span(page_idx, bbox)
            if crop is None or getattr(crop, "size", None) in (None, (0, 0)):
                continue
            width, height = crop.size
            if width < 5 or height < 5:
                continue
            with torch.no_grad():
                res = latex_ocr(crop.convert("RGB"))
                if res:
                    pix2tex_results[i] = str(res).strip()
        except Exception:
            pass
        if progress_cb and total:
            progress_cb((i + 1) / total)
        if i % 10 == 0:
            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

    interline_queue = [r for (_, _, st_), r in zip(equation_spans, pix2tex_results) if st_ == "interline_equation"]
    inline_queue = [r for (_, _, st_), r in zip(equation_spans, pix2tex_results) if st_ != "interline_equation"]

    def _fill_blank_display_eqs(text, queue):
        out, pos, idx = [], 0, 0
        for m in re.finditer(r"\$\$(.*?)\$\$", text, re.DOTALL):
            out.append(text[pos:m.start()])
            inner = m.group(1).strip()
            if not inner and idx < len(queue) and queue[idx]:
                out.append(f"$${queue[idx]}$$")
            else:
                out.append(m.group(0))
            idx += 1
            pos = m.end()
        out.append(text[pos:])
        return "".join(out)

    def _fill_blank_inline_eqs(text, queue):
        out, pos, idx = [], 0, 0
        for m in re.finditer(r"(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)", text, re.DOTALL):
            out.append(text[pos:m.start()])
            inner = m.group(1).strip()
            if not inner and idx < len(queue) and queue[idx]:
                out.append(f"${queue[idx]}$")
            else:
                out.append(m.group(0))
            idx += 1
            pos = m.end()
        out.append(text[pos:])
        return "".join(out)

    filled = sum(1 for r in pix2tex_results if r)
    md_text = _fill_blank_display_eqs(md_text, interline_queue)
    md_text = _fill_blank_inline_eqs(md_text, inline_queue)
    return md_text, filled


def run_paddle_hindi_pass(md_text, text_spans, renderer, paddle_hi, progress_cb=None):
    hindi_fix_count = 0
    total = len(text_spans)
    errors = []

    for i, (page_idx, bbox, mineru_text) in enumerate(text_spans):
        try:
            if not mineru_text or _DEVANAGARI_RE.search(mineru_text):
                continue
            crop = renderer.crop_span(page_idx, bbox)
            if crop is None:
                continue
            crop_np = np.array(crop.convert("RGB"))
            try:
                res = paddle_hi.predict(crop_np)
                paddle_text = " ".join(t for item in res for t in item.get("rec_texts", [])) if res else ""
            except AttributeError:
                res = paddle_hi.ocr(crop_np, cls=True)
                paddle_text = " ".join(line[1][0] for block in (res or []) for line in (block or []))

            if paddle_text and _DEVANAGARI_RE.search(paddle_text) and mineru_text in md_text:
                md_text = md_text.replace(mineru_text, paddle_text, 1)
                hindi_fix_count += 1
        except Exception as e:
            errors.append(str(e))
        if progress_cb and total:
            progress_cb((i + 1) / total)

    return md_text, hindi_fix_count, errors


def clean_markdown_content(text):
    """MCQ-formatting safety net: vertical A/B/C/D options, question numbers
    made bold plain-text (not pandoc auto-numbered lists) so PDF's real
    question numbers survive untouched into the docx."""
    logical_lines = []
    for line in text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        bullet_match = re.match(r"^[-*+]\s+(.*)$", stripped)
        content = bullet_match.group(1) if bullet_match else stripped

        qnum_match = _QUESTION_NUM_RE.match(content)
        if qnum_match and not re.match(r"^\([A-Ea-e]\)", qnum_match.group(2).strip()):
            content = f"**{qnum_match.group(1)}.** {qnum_match.group(2)}"

        markers = list(_OPTION_MARKER_RE.finditer(content))
        if len(markers) >= 2:
            first_start = markers[0].start()
            leading_text = content[:first_start].strip()
            if leading_text:
                logical_lines.append((True, f"(A) {leading_text}"))
            for i, m in enumerate(markers):
                start = m.start()
                end = markers[i + 1].start() if i + 1 < len(markers) else len(content)
                logical_lines.append((True, content[start:end].strip()))
        else:
            is_single_option = bool(re.match(r"^\([A-Ea-e]\)", content))
            logical_lines.append((is_single_option, content))

    blocks, current = [], []

    def flush():
        if current:
            blocks.append("  \n".join(current))
            current.clear()

    for is_option, content in logical_lines:
        if is_option:
            current.append(content)
        else:
            flush()
            blocks.append(content)
    flush()
    return "\n\n".join(blocks)


def run_pandoc(md_dir, md_filename, docx_name):
    pandoc_cmd = [
        "pandoc", md_filename,
        "-f", "markdown+tex_math_dollars+tex_math_single_backslash+raw_tex+pipe_tables+grid_tables+raw_html",
        "-o", docx_name,
    ]
    result = subprocess.run(pandoc_cmd, cwd=md_dir, capture_output=True, text=True)
    return result


def looks_like_plain_text(latex_src):
    """True if this is really plain prose (e.g. an answer-label like
    'Ans. (D)') that MinerU mis-tagged as an equation span, not real math."""
    src = latex_src.strip()
    if not src:
        return True
    has_math_symbols = bool(re.search(r"[\\^_=+\-*/{}]", src))
    return (not has_math_symbols) and bool(_PLAIN_TEXT_RE.match(src))


def latex_to_omath(latex_src):
    import latex2mathml.converter as l2m
    import mathml2omml
    from lxml import etree

    mathml = l2m.convert(latex_src.strip())
    omml_str = mathml2omml.convert(mathml)
    wrapped = f'<m:root xmlns:m="{_M_NS}">{omml_str}</m:root>'
    return etree.fromstring(wrapped.encode("utf-8"))[0]


def repair_docx_equations(docx_in, docx_out):
    from lxml import etree

    tmp_dir = docx_out + "_tmpext"
    shutil.rmtree(tmp_dir, ignore_errors=True)
    with zipfile.ZipFile(docx_in) as z:
        z.extractall(tmp_dir)
    doc_path = os.path.join(tmp_dir, "word", "document.xml")
    tree = etree.parse(doc_path)
    root = tree.getroot()
    fixed, failed, skipped_plain = 0, 0, 0
    for t_el in root.iter(f"{{{_W_NS}}}t"):
        text = (t_el.text or "").strip()
        m = re.fullmatch(r"\$\$(.+?)\$\$", text, re.DOTALL) or re.fullmatch(r"\$(.+?)\$", text, re.DOTALL)
        if not m:
            continue
        latex_src = next(g for g in m.groups() if g is not None)

        if looks_like_plain_text(latex_src):
            t_el.text = latex_src.strip()
            skipped_plain += 1
            continue

        try:
            omath_el = latex_to_omath(latex_src)
        except Exception:
            failed += 1
            continue
        run_el = t_el.getparent()
        run_el.getparent().replace(run_el, omath_el)
        fixed += 1
    tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(docx_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for base, _, fs in os.walk(tmp_dir):
            for fn in fs:
                full = os.path.join(base, fn)
                zout.write(full, os.path.relpath(full, tmp_dir))
    shutil.rmtree(tmp_dir, ignore_errors=True)
    return fixed, skipped_plain, failed


# ----------------------------------------------------------------------
# Sidebar options
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Settings")
    do_pix2tex = st.checkbox("Equation double-check (pix2tex)", value=True)
    do_paddle = st.checkbox("Hindi double-check (PaddleOCR)", value=True)
    st.caption(
        "Note: MinerU / pix2tex / PaddleOCR bahut heavy hain (GPU + model "
        "downloads chahiye). Yeh app usually local GPU machine par chalao, "
        "Streamlit Cloud jaisi free hosting par nahi chalega."
    )

_missing_deps = check_system_deps()
if _missing_deps:
    st.error(
        "⚠️ Yeh system binaries missing hain: **" + ", ".join(_missing_deps) + "**.\n\n"
        "- `mineru` missing ho to: `pip install -U \"mineru[all]\"` chalao aur confirm karo "
        "ki `mineru --help` terminal mein chalta hai.\n"
        "- `pandoc` / `pdftoppm` missing ho to: yeh apt/system packages hain — "
        "`sudo apt-get install poppler-utils pandoc` (ya deployment platform ka "
        "`packages.txt` mechanism, jo Streamlit Community Cloud use karta hai).\n\n"
        "**Agar tum Streamlit Community Cloud par ho:** iska free tier is pipeline "
        "(MinerU + pix2tex + PaddleOCR ke multi-GB models, GPU-preferred workload) "
        "ke liye resource-wise fit nahi hai — RAM/disk/build-time limits hit ho "
        "jaayenge. Isko apne GPU wale VM/server par self-host karna, ya Hugging Face "
        "Spaces (GPU tier) jaisi platform use karna better rahega. Details README.md mein hain."
    )
    st.stop()

# ----------------------------------------------------------------------
# Main flow
# ----------------------------------------------------------------------
uploaded_file = st.file_uploader("Apni PDF upload karo", type=["pdf"])

if uploaded_file is not None:
    st.success(f"Upload ho gayi: {uploaded_file.name}")

    if st.button("🚀 Process PDF", type="primary"):
        work_dir = tempfile.mkdtemp(prefix="pdf2word_")
        pdf_path = os.path.join(work_dir, "doc_input.pdf")
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        mineru_out_dir = os.path.join(work_dir, "mineru_out")
        os.makedirs(mineru_out_dir, exist_ok=True)

        status = st.status("Pipeline chal raha hai...", expanded=True)

        try:
            # --- Step 1: Hindi detection --------------------------------
            status.write("🔎 Devanagari ratio check ho raha hai...")
            use_devanagari_mode, ratio = pdf_has_significant_hindi(pdf_path)
            status.write(f"Devanagari character ratio: {ratio:.2%} → mode: "
                         f"**{'devanagari' if use_devanagari_mode else 'default/english'}**")

            # --- Step 2: Run MinerU --------------------------------------
            status.write("🧠 MinerU chal raha hai (pehli baar models HuggingFace se download honge)...")
            cmd, result = run_mineru(pdf_path, mineru_out_dir, use_devanagari_mode)
            status.write(f"Command: `{' '.join(cmd)}`")
            if result.returncode != 0:
                status.update(label="MinerU fail ho gaya", state="error")
                st.error("MinerU fail ho gaya. Neeche poora error dekho:")
                st.code((result.stdout or "")[-3000:] + "\n" + (result.stderr or "")[-3000:])
                st.stop()

            md_candidates, middle_json_candidates = locate_mineru_outputs(mineru_out_dir)
            if not md_candidates:
                status.update(label="MinerU ne .md file nahi banayi", state="error")
                st.error(f"MinerU ne koi .md file nahi banayi. {mineru_out_dir} check karo.")
                st.stop()

            md_file_path = md_candidates[0]
            md_dir = os.path.dirname(md_file_path)
            with open(md_file_path, "r", encoding="utf-8") as f:
                md_text = f.read()
            status.write(f"✅ MinerU markdown mila ({len(md_text.split())} words)")

            middle_json_data = None
            if middle_json_candidates:
                with open(middle_json_candidates[0], "r", encoding="utf-8") as f:
                    middle_json_data = json.load(f)
                status.write("✅ middle.json mila")
            else:
                status.write("⚠️ [skip] middle.json nahi mila — equation/Hindi double-check skip honge, baseline MinerU output use hoga.")

            # --- Step 3: Collect spans -----------------------------------
            equation_spans, text_spans, page_dims = [], [], {}
            if middle_json_data:
                try:
                    equation_spans, text_spans, page_dims = collect_spans(middle_json_data)
                    status.write(f"📐 {len(equation_spans)} equation span(s), {len(text_spans)} text span(s) mile.")
                except Exception as e:
                    status.write(f"⚠️ [skip] middle.json parse fail: {e}")

            renderer = None
            if equation_spans or text_spans:
                renderer = PageRenderer(pdf_path, page_dims)

            # --- Step 4: pix2tex pass ------------------------------------
            if do_pix2tex and equation_spans and renderer is not None:
                status.write("🧮 Equations ko pix2tex se verify kar rahe hain...")
                latex_ocr = load_pix2tex()
                bar = st.progress(0.0)
                md_text, filled = run_pix2tex_pass(
                    md_text, equation_spans, renderer, latex_ocr,
                    progress_cb=lambda p: bar.progress(p),
                )
                bar.empty()
                status.write(f"✅ {filled}/{len(equation_spans)} equation(s) pix2tex se recognize hui.")
            else:
                status.write("⏭️ pix2tex pass skip kiya.")

            # --- Step 5: PaddleOCR Hindi pass -----------------------------
            if do_paddle and text_spans and renderer is not None:
                status.write("🇮🇳 Hindi text PaddleOCR se cross-check kar rahe hain...")
                paddle_hi = load_paddle_hi()
                bar2 = st.progress(0.0)
                md_text, hindi_fix_count, errs = run_paddle_hindi_pass(
                    md_text, text_spans, renderer, paddle_hi,
                    progress_cb=lambda p: bar2.progress(p),
                )
                bar2.empty()
                status.write(f"✅ {hindi_fix_count} Hindi text span(s) PaddleOCR se fix ki gayin.")
                if errs:
                    status.write(f"⚠️ {len(errs)} span(s) mein check fail hui (ignore kiya gaya).")
            else:
                status.write("⏭️ PaddleOCR Hindi pass skip kiya.")

            # --- Step 6: MCQ / structure cleanup --------------------------
            status.write("🧹 MCQ/structure cleanup ho raha hai...")
            md_text = clean_markdown_content(md_text)
            clean_md_name = "output_clean.md"
            with open(os.path.join(md_dir, clean_md_name), "w", encoding="utf-8") as f:
                f.write(md_text)

            # --- Step 7: pandoc md -> docx ---------------------------------
            status.write("📝 Pandoc se Markdown → .docx bana rahe hain...")
            base_name = os.path.splitext(uploaded_file.name)[0]
            docx_name = f"{base_name}.docx"
            result = run_pandoc(md_dir, clean_md_name, docx_name)
            docx_path = os.path.join(md_dir, docx_name)
            if result.returncode != 0 or not os.path.exists(docx_path):
                status.update(label="Pandoc fail ho gaya", state="error")
                st.error("Pandoc fail ho gaya:")
                st.code((result.stdout or "") + "\n" + (result.stderr or ""))
                st.stop()
            if result.stderr.strip():
                with st.expander("Pandoc warnings (equations jo TeX se convert nahi ho payin)"):
                    st.code(result.stderr)
            status.write("✅ Word document ban gaya.")

            # --- Step 8: OMML equation repair ------------------------------
            status.write("🔧 Equation-repair pass (real OMML equations bana rahe hain)...")
            repaired_path = docx_path.replace(".docx", "_final.docx")
            fixed, skipped_plain, failed = repair_docx_equations(docx_path, repaired_path)
            final_path = repaired_path if os.path.exists(repaired_path) else docx_path
            status.write(
                f"✅ Equation repair: {fixed} fixed, {skipped_plain} plain-text "
                f"(galat-tagged) chhode gaye, {failed} manually check karni padegi."
            )

            status.update(label="✅ Pipeline complete!", state="complete")

            with open(final_path, "rb") as f:
                docx_bytes = f.read()

            st.session_state["final_docx_bytes"] = docx_bytes
            st.session_state["final_docx_name"] = os.path.basename(final_path)

        except Exception as e:
            status.update(label="❌ Error aaya", state="error")
            st.exception(e)

if "final_docx_bytes" in st.session_state:
    st.download_button(
        "⬇️ Download Word Document",
        data=st.session_state["final_docx_bytes"],
        file_name=st.session_state["final_docx_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
