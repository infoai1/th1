from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    None: "left",  # Word often omits explicit alignment; treat as left
}

def _pt(font_size):
    try:
        return font_size.pt if font_size else None
    except Exception:
        return None

def _any_bold(paragraph):
    return any(r.bold for r in paragraph.runs if r.text)

def _bold_fraction(paragraph):
    runs = paragraph.runs
    if not runs:
        return 0.0
    total = sum(len(r.text or "") for r in runs)
    if total == 0:
        return 0.0
    bold_chars = sum(len(r.text or "") for r in runs if r.bold)
    return bold_chars / max(total, 1)

def _max_font_size(paragraph):
    sizes = [_pt(r.font.size) for r in paragraph.runs if r.font is not None]
    sizes = [s for s in sizes if s is not None]
    return max(sizes) if sizes else None

def _avg_font_size(paragraph):
    sizes = [_pt(r.font.size) for r in paragraph.runs if r.font is not None and r.font.size is not None]
    if not sizes:
        return None
    return sum(sizes) / len(sizes)

def _style_name(paragraph):
    try:
        return (paragraph.style.name or "").lower()
    except Exception:
        return ""

def _looks_sentence_like(text):
    t = (text or "").strip()
    if not t:
        return True
    # Ends with punctuation or quote; contains verbs/punctuation typical of sentences
    end_punct = t.endswith(".") or t.endswith("!") or t.endswith("?") or t.endswith('"') or t.endswith("'")
    multi = t.count(".") + t.count("!") + t.count("?") >= 1
    longish = len(t.split()) >= 10
    return end_punct or multi or longish

def _is_quoted_oneliner(text):
    t = (text or "").strip()
    if len(t.split()) <= 12 and ((t.startswith('"') and t.endswith('"')) or (t.startswith("'") and t.endswith("'"))):
        return True
    return False

def _alignment(paragraph):
    return ALIGN_MAP.get(paragraph.alignment, "left")

def classify_levels(paragraph, text, rules):
    # Returns dict {"is_h1": bool, "is_h2": bool, "is_h3": bool, "score": int, "features": {...}}
    words = text.split()
    word_count = len(words)
    all_caps = text.isupper() and any(c.isalpha() for c in text)
    short_phrase = word_count <= rules.get("max_header_words", 15) and len(text) <= 120
    avg_size = _avg_font_size(paragraph)
    max_size = _max_font_size(paragraph)
    bold_frac = _bold_fraction(paragraph)
    any_bold = _any_bold(paragraph)
    align = _alignment(paragraph)
    style = _style_name(paragraph)
    sentence_like = _looks_sentence_like(text) if rules.get("suppress_sentences", True) else False
    quoted_oneliner = _is_quoted_oneliner(text) if rules.get("suppress_quotes", True) else False

    # Style hints override (Heading 1/2/3)
    style_h1 = "heading 1" in style or style.strip() == "heading1"
    style_h2 = "heading 2" in style or style.strip() == "heading2"
    style_h3 = "heading 3" in style or style.strip() == "heading3"

    def level_match(level_key):
        lvl = rules["levels"][level_key]
        if not lvl.get("enabled", True):
            return False, 0

        size_ok = (avg_size and avg_size >= lvl["min_size"]) or (max_size and max_size >= lvl["min_size"])
        if not size_ok:
            return False, 0

        align_ok = (align in lvl.get("allowed_align", ["left", "center", "right"]))
        if not align_ok:
            return False, 0

        if lvl.get("require_bold", False) and not any_bold and bold_frac < 0.4:
            return False, 0

        if lvl.get("require_short_phrase", True) and not short_phrase:
            return False, 0

        # Score accumulation
        score = 0
        if any_bold or bold_frac >= 0.6:
            score += 1
        if all_caps:
            score += 1
        if align == "center":
            score += 1
        if short_phrase:
            score += 1
        if "heading" in style:
            score += 2

        # Penalties
        if sentence_like:
            score -= 2
        if quoted_oneliner:
            score -= 2

        return True, score

    # Try style overrides first
    is_h1 = style_h1
    is_h2 = style_h2
    is_h3 = style_h3
    score = 0

    # If style didn't decide, use rules
    if not any([is_h1, is_h2, is_h3]):
        for key in ["h1", "h2", "h3"]:
            ok, sc = level_match(key)
            if ok:
                if key == "h1": is_h1 = True
                if key == "h2": is_h2 = True
                if key == "h3": is_h3 = True
                score = max(score, sc)

    features = {
        "all_caps": all_caps,
        "short_phrase": short_phrase,
        "avg_font_size": round(avg_size, 2) if avg_size else None,
        "max_font_size": round(max_size, 2) if max_size else None,
        "bold_fraction": round(bold_frac, 2),
        "any_bold": any_bold,
        "align": align,
        "style": style,
        "sentence_like": sentence_like,
        "quoted_oneliner": quoted_oneliner,
        "word_count": word_count
    }

    # Overall is_header if any level matched
    is_header = any([is_h1, is_h2, is_h3])
    return {"is_h1": is_h1, "is_h2": is_h2, "is_h3": is_h3, "is_header": is_header, "score": score, "features": features}

def parse_docx(docx_file, rules):
    doc = Document(docx_file)
    rows = []
    for idx, para in enumerate(doc.paragraphs):
        text = (para.text or "").strip()
        if not text:
            continue
        c = classify_levels(para, text, rules)
        feat = c["features"]
        rows.append({
            "idx": idx,
            "text": text,
            "is_h1": c["is_h1"],
            "is_h2": c["is_h2"],
            "is_h3": c["is_h3"],
            "is_header": c["is_header"],
            "score": c["score"],
            "all_caps": feat["all_caps"],
            "short_phrase": feat["short_phrase"],
            "avg_font_size": feat["avg_font_size"],
            "max_font_size": feat["max_font_size"],
            "bold_fraction": feat["bold_fraction"],
            "any_bold": feat["any_bold"],
            "align": feat["align"],
            "style": feat["style"],
            "sentence_like": feat["sentence_like"],
            "quoted_oneliner": feat["quoted_oneliner"],
            "word_count": feat["word_count"],
        })
    return rows
