# PDF → Word (MinerU + pix2tex + PaddleOCR) — Streamlit App

Original Colab notebook (`pdf_to_word_mineru_pix2tex_paddleocr_FIXED.ipynb`) ka
Streamlit version. Same 3-tool pipeline:

1. **MinerU** (CLI) — primary layout/text/table/equation extraction (`-l devanagari` mode Hindi ke liye).
2. **pix2tex** — equation crops ka second-opinion LaTeX recognition.
3. **PaddleOCR (hi)** — Devanagari text crops ka second-opinion recognition.
4. **pandoc** — cleaned Markdown → `.docx`.
5. **OMML repair pass** — jo bhi TeX pandoc khud na convert kar paya, use `latex2mathml` + `mathml2omml` se real Word equations bana deta hai.

## ⚠️ Important — yeh app GPU machine par chalao

MinerU + pix2tex + PaddleOCR sab heavy models hain jo pehli baar HuggingFace
se download hote hain aur GPU par best chalte hain. **Streamlit Community
Cloud jaisi free/shared hosting par yeh nahi chalega** — apne GPU wale
server/desktop/Colab-VM par local run karo.

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
