import pandas as pd

def _yield_chunks(words, min_words=200, max_words=250, overlap=0.2):
    if max_words < min_words:
        max_words = min_words
    if overlap < 0 or overlap >= 1:
        overlap = 0.0

    step = max(1, int(max_words * (1 - overlap)))
    i = 0
    while i < len(words):
        chunk = words[i:i+max_words]
        if len(chunk) >= min_words or (i == 0 and len(words) > 0):
            yield " ".join(chunk)
        i += step

def build_csv_rows(rows, book_name, author_name, min_words, max_words, overlap):
    csv_rows = []
    current = {"h1": None, "h2": None, "h3": None}
    body_accum = []

    def flush_body():
        nonlocal body_accum, current
        if not body_accum:
            return
        text = " ".join(body_accum).strip()
        body_accum = []
        if not text:
            return
        words = text.split()
        chapter_parts = [p for p in [current["h1"], current["h2"], current["h3"]] if p]
        chapter_name = " | ".join(chapter_parts) if chapter_parts else "Introduction"
        for chunk in _yield_chunks(words, min_words=min_words, max_words=max_words, overlap=overlap):
            csv_rows.append({
                "book_name": book_name or "Unknown Book",
                "author_name": author_name or "Unknown Author",
                "chapter_name": chapter_name,
                "text_chunk": chunk
            })

    for row in rows:
        # Determine the strongest level claimed by this row
        level = None
        if row.get("is_h1"):
            level = "h1"
        elif row.get("is_h2"):
            level = "h2"
        elif row.get("is_h3"):
            level = "h3"

        if level:  # header row
            flush_body()
            if level == "h1":
                current["h1"] = row["text"]
                current["h2"] = None
                current["h3"] = None
            elif level == "h2":
                current["h2"] = row["text"]
                current["h3"] = None
            elif level == "h3":
                current["h3"] = row["text"]
        else:
            body_accum.append(row["text"])

    flush_body()
    return pd.DataFrame(csv_rows)
