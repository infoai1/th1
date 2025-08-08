import streamlit as st
import pandas as pd
from io import BytesIO
from header_utils import parse_docx
from chunker import build_csv_rows

st.set_page_config(page_title="DOCX → CSV Chapter Chunker v2", layout="wide")
st.title("📚 DOCX → CSV Chapter Chunker (v2)")
st.caption("Multi-level headers (H1/H2/H3), alignment filters, bold requirements, sentence/quote suppression, chunk & export CSV.")

with st.expander("ℹ️ Quick help", expanded=True):
    st.markdown(
        "- Default thresholds match your file: **H1 ≥ 14pt (bold)**, **H2 ≥ 13pt**, **H3 ≥ 13pt**, body ≈ 13pt.\n"
        "- Use **Preview & Edit** to toggle any row's header level.\n"
        "- Misdetected quotes? Keep **Suppress sentence/quoted lines** enabled."
    )

colBA1, colBA2 = st.columns(2)
book_name = colBA1.text_input("Book Name", value="Spirit of Islam")
author_name = colBA2.text_input("Author Name", value="Unknown Author")

st.subheader("Header Detection Rules")
st.markdown("Configure per-level thresholds and alignment. Not all headers need to be ALL CAPS.")

with st.container():
    c1, c2, c3, c4 = st.columns(4)
    max_header_words = c1.slider("Max header words", 3, 20, 15)
    suppress_sentences = c2.checkbox("Suppress sentence-like lines", value=True)
    suppress_quotes = c3.checkbox("Suppress quoted one-liners", value=True)
    auto_detect = c4.checkbox("Enable Auto-detect", value=True)

st.markdown("**H1 (Main chapter)**")
h1c1, h1c2, h1c3, h1c4 = st.columns(4)
h1_enabled = h1c1.checkbox("Enable H1", value=True)
h1_min_size = h1c2.number_input("H1 min font size (pt)", value=14, step=1)
h1_require_bold = h1c3.checkbox("H1 require bold", value=True)
h1_align = h1c4.multiselect("H1 align allowed", ["left", "center", "right"], default=["left", "center", "right"])

st.markdown("**H2**")
h2c1, h2c2, h2c3, h2c4 = st.columns(4)
h2_enabled = h2c1.checkbox("Enable H2", value=True)
h2_min_size = h2c2.number_input("H2 min font size (pt)", value=13, step=1)
h2_require_bold = h2c3.checkbox("H2 require bold", value=False)
h2_align = h2c4.multiselect("H2 align allowed", ["left", "center", "right"], default=["left", "center", "right"])

st.markdown("**H3**")
h3c1, h3c2, h3c3, h3c4 = st.columns(4)
h3_enabled = h3c1.checkbox("Enable H3", value=True)
h3_min_size = h3c2.number_input("H3 min font size (pt)", value=13, step=1)
h3_require_bold = h3c3.checkbox("H3 require bold", value=False)
h3_align = h3c4.multiselect("H3 align allowed", ["left", "center", "right"], default=["left", "center", "right"])

st.subheader("Chunk Settings")
cc1, cc2, cc3 = st.columns(3)
min_words = cc1.slider("Min words per chunk", 50, 500, 200, step=10)
max_words = cc2.slider("Max words per chunk", 80, 800, 250, step=10)
overlap_pct = cc3.slider("Overlap (%)", 0, 60, 20, step=5)
overlap = overlap_pct / 100.0

uploaded = st.file_uploader("Upload a DOCX file", type=["docx"])

def build_rules():
    return {
        "auto_detect": auto_detect,
        "max_header_words": max_header_words,
        "suppress_sentences": suppress_sentences,
        "suppress_quotes": suppress_quotes,
        "levels": {
            "h1": {"enabled": h1_enabled, "min_size": float(h1_min_size), "require_bold": bool(h1_require_bold),
                   "allowed_align": [a.lower() for a in h1_align], "require_short_phrase": True},
            "h2": {"enabled": h2_enabled, "min_size": float(h2_min_size), "require_bold": bool(h2_require_bold),
                   "allowed_align": [a.lower() for a in h2_align], "require_short_phrase": True},
            "h3": {"enabled": h3_enabled, "min_size": float(h3_min_size), "require_bold": bool(h3_require_bold),
                   "allowed_align": [a.lower() for a in h3_align], "require_short_phrase": True},
        }
    }

if uploaded is not None and st.button("Preview & Edit Headers", type="primary"):
    try:
        rules = build_rules()
        rows = parse_docx(uploaded, rules) if auto_detect else []
        if not auto_detect:
            st.warning("Auto-detect disabled: start with no headers, then manually mark rows below.")
            # Build rows with empty detection but include text content
            from docx import Document
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            doc = Document(uploaded)
            idx = 0
            tmp = []
            for p in doc.paragraphs:
                text = (p.text or "").strip()
                if not text:
                    continue
                tmp.append({
                    "idx": idx, "text": text,
                    "is_h1": False, "is_h2": False, "is_h3": False, "is_header": False,
                    "score": 0, "all_caps": text.isupper(),
                    "short_phrase": len(text.split()) <= max_header_words,
                    "avg_font_size": None, "max_font_size": None, "bold_fraction": 0.0,
                    "any_bold": False, "align": "left", "style": "", "sentence_like": False,
                    "quoted_oneliner": False, "word_count": len(text.split())
                })
                idx += 1
            rows = tmp
        st.session_state["rows"] = rows
    except Exception as e:
        st.error(f"Failed to parse DOCX: {e}")

if "rows" in st.session_state:
    st.subheader("Review detection (edit H1/H2/H3 flags if needed)")
    df = pd.DataFrame(st.session_state["rows"])
    cols = ["idx","text","is_h1","is_h2","is_h3","score","align","any_bold","avg_font_size","max_font_size","all_caps","short_phrase","word_count"]
    df = df[cols]

    edited = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "text": st.column_config.TextColumn(width="large"),
            "is_h1": st.column_config.CheckboxColumn(help="Mark as Heading 1"),
            "is_h2": st.column_config.CheckboxColumn(help="Mark as Heading 2"),
            "is_h3": st.column_config.CheckboxColumn(help="Mark as Heading 3"),
            "score": st.column_config.NumberColumn(format="%.0f"),
            "avg_font_size": st.column_config.NumberColumn(format="%.1f"),
            "max_font_size": st.column_config.NumberColumn(format="%.1f"),
        },
        disabled=["idx","score","align","any_bold","avg_font_size","max_font_size","all_caps","short_phrase","word_count","text"]
    )

    st.markdown("**Current headers preview:**")
    preview = edited[(edited["is_h1"]) | (edited["is_h2"]) | (edited["is_h3"])][["idx","is_h1","is_h2","is_h3","text"]].head(50)
    st.dataframe(preview, use_container_width=True)

    if st.button("Apply Edits"):
        merged_h1 = edited.set_index("idx")["is_h1"].to_dict()
        merged_h2 = edited.set_index("idx")["is_h2"].to_dict()
        merged_h3 = edited.set_index("idx")["is_h3"].to_dict()
        for r in st.session_state["rows"]:
            idx = r["idx"]
            r["is_h1"] = bool(merged_h1.get(idx, r["is_h1"]))
            r["is_h2"] = bool(merged_h2.get(idx, r["is_h2"]))
            r["is_h3"] = bool(merged_h3.get(idx, r["is_h3"]))
            r["is_header"] = r["is_h1"] or r["is_h2"] or r["is_h3"]
        st.success("Edits applied.")

    if st.button("Generate CSV"):
        final_rows = st.session_state["rows"]
        out_df = build_csv_rows(final_rows, book_name, author_name, min_words, max_words, overlap)
        if out_df.empty:
            st.warning("No content produced. Try lowering min words or check header detection.")
        else:
            st.write(out_df.head(5))
            buf = BytesIO()
            out_df.to_csv(buf, index=False, encoding="utf-8-sig")
            buf.seek(0)
            st.download_button(
                "⬇️ Download CSV",
                data=buf,
                file_name="output.csv",
                mime="text/csv"
            )
            st.success(f"CSV ready with {len(out_df)} chunks.")
