# PDF → Word (MinerU + pix2tex + PaddleOCR) — Streamlit App

Original Colab notebook (`pdf_to_word_mineru_pix2tex_paddleocr_FIXED.ipynb`) ka
Streamlit version. Same 3-tool pipeline:

1. **MinerU** (CLI) — primary layout/text/table/equation extraction (`-l devanagari` mode Hindi ke liye).
2. **pix2tex** — equation crops ka second-opinion LaTeX recognition.
3. **PaddleOCR (hi)** — Devanagari text crops ka second-opinion recognition.
4. **pandoc** — cleaned Markdown → `.docx`.
5. **OMML repair pass** — jo bhi TeX pandoc khud na convert kar paya, use `latex2mathml` + `mathml2omml` se real Word equations bana deta hai.

## ⚠️ Important — Streamlit Community Cloud par yeh (reliably) nahi chalega

MinerU + pix2tex + PaddleOCR sab heavy models hain (multi-GB, HuggingFace se
pehli baar download hote hain) aur GPU par best chalte hain. Streamlit
Community Cloud (free tier) mein:

- No GPU (CPU-only — MinerU/pix2tex/PaddleOCR bahut slow ya OOM ho sakte hain)
- Limited RAM (~1 GB typically) aur ephemeral/limited disk
- Build/deploy timeouts jo `mineru[all]` jaisi heavy install ko fail kar sakte hain
- `mineru` command PATH par install ho bhi jaaye, resource limits ke wajah se
  pipeline crash/hang ho sakti hai

Agar tumhe `FileNotFoundError: [Errno 2] No such file or directory: 'mineru'`
jaisa error mil raha hai Streamlit Cloud par, iska matlab hai `mineru[all]`
ka pip install us build mein fail/skip ho gaya (build logs check karo) — aur
chahe woh install ho bhi jaaye, poora pipeline us free tier par practically
nahi chalega.

**Recommended**: apne GPU wale server/desktop/VM par local run karo (neeche
Setup dekho), ya self-hosted Docker container, ya Hugging Face Spaces
(GPU tier) par deploy karo.

`packages.txt` file is repo mein already included hai (Streamlit Community
Cloud isse padh kar `poppler-utils` aur `pandoc` apt se install karta hai) —
lekin yeh sirf poppler/pandoc ke liye hai, `mineru`/`torch`/`paddleocr` jaisi
heavy Python deps ke resource-requirement ka fix nahi.

## Setup (Ubuntu/Debian + NVIDIA GPU maan kar)

```bash
# System packages
sudo apt-get update && sudo apt-get install -y poppler-utils pandoc

# Python env
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip

# Core app + light deps
pip install streamlit numpy pypdf "pypdfium2==4.30.0" lxml latex2mathml mathml2omml

# MinerU (pulls the "pipeline" backend — works on a single T4-class GPU)
pip install -U "mineru[all]"

# pix2tex (equation OCR)
pip install "pix2tex[gui]"

# PaddleOCR (Hindi) — GPU build (adjust URL for your CUDA version if needed)
pip install paddlepaddle-gpu -f https://www.paddlepaddle.org.cn/whl/linux/mkl/avx/stable.html
pip install paddleocr
```

If you don't have a GPU, replace `paddlepaddle-gpu` with plain `paddlepaddle`
and expect MinerU/pix2tex to run (slower) on CPU.

## Run

```bash
streamlit run app.py
```

Browser mein `http://localhost:8501` khul jayega. PDF upload karo, sidebar
mein pix2tex/PaddleOCR passes on/off karo, **Process PDF** dabao, aur end
mein **Download Word Document** button se final `.docx` mil jayega.

## Notes / gotchas (from the original notebook's bug history)

- Models (`pix2tex`, `paddleocr`) `st.cache_resource` se cached hain, taaki
  har PDF ke liye dobara load na hon — sirf Streamlit server restart par
  reload honge.
- `middle.json` ka exact field-naming MinerU version ke hisaab se thoda
  badal sakta hai — is wajah se span-collection defensively try/except mein
  wrapped hai; agar spans 0 aayen to app status log mein dikh jayega.
- Question numbers (`1.`, `11.`...) bold plain-text banaye jaate hain
  (ordered-list nahi), taaki Word auto-renumber na kare aur PDF ka asli
  number safe rahe.
- Equation spans jo asal mein plain text hain (jaise `Ans. (D)`) unhe OMML
  mein force nahi kiya jaata — seedha plain text ki tarah rakha jaata hai.
