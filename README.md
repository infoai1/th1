# DOCX ➜ CSV Chapter Chunker (Streamlit) — v2

✔ Multi-level headers (H1 / H2 / H3)
✔ Alignment filters (Left / Center / Right)
✔ "Require Bold" per level (H1 defaults to **bold**)
✔ Font thresholds: H1 ≥ 14 pt, H2 ≥ 13 pt, H3 ≥ 13 pt (tunable)
✔ Downweight sentence-like/quoted one-liners to avoid false headers
✔ Edit header levels inline before exporting CSV
✔ Chapter label is composed as `H1 | H2 | H3`

## Deploy (No local install)

### Option A — Hugging Face Spaces (Easy)
1) Create a free account at huggingface.co
2) New Space → Type: *Streamlit* → Runtime: *Python*
3) Upload these files: `app.py`, `header_utils.py`, `chunker.py`, `requirements.txt`, `README.md`

### Option B — Streamlit Community Cloud
1) Put the files into a GitHub repo
2) Go to share.streamlit.io → New app → pick repo → main file `app.py`

## Use
- Upload `.docx`, set **Book/Author**.
- Set font thresholds (default: H1=14, H2=13, H3=13) and **Require Bold** for H1 if your main chapter header is bold.
- Choose allowed alignments for each level.
- Preview & edit, then Generate CSV.

## Tips
- If quotes were misdetected, keep **Suppress sentence-like/quoted lines** enabled.
- Not all headers must be ALL CAPS; the detector uses multiple signals.
