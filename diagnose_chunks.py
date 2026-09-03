"""
Diagnostic script: Dumps all chunks with character analysis.
Run this to see exactly what text is in each chunk and spot
problematic characters/patterns.
"""

import re
import unicodedata
from docx import Document


def extract_text_from_docx(filepath):
    doc = Document(filepath)
    texts = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return "\n".join(texts)


def clean_ocr_text(text):
    text = re.sub(r"(?i)(www\.[a-z0-9\-]+\.gov\.np|lawcommission\.gov\.np)", "", text)
    cleaned_lines = []
    for line in text.split("\n"):
        line = re.sub(r"[ \t]+", " ", line).strip()
        if not line:
            continue
        if re.fullmatch(r"[0-9०-९]+", line):
            continue
        has_devanagari = bool(re.search(r"[\u0900-\u097F]", line))
        if not has_devanagari:
            if not re.search(r"(?i)\b(schedule|part|section|article)\b", line):
                if len(line) < 10 and not re.search(r"[a-zA-Z]{3,}", line):
                    continue
        line = re.sub(r"^\|\s*", "", line)
        cleaned_lines.append(line)
    return "\n".join(cleaned_lines)


def analyze_char(ch):
    """Return character info: codepoint, name, category."""
    cp = f"U+{ord(ch):04X}"
    try:
        name = unicodedata.name(ch)
    except ValueError:
        name = "<UNNAMED>"
    cat = unicodedata.category(ch)
    return cp, name, cat


def analyze_text(text):
    """Analyze a chunk of text for unusual characters."""
    issues = []
    char_freq = {}

    for i, ch in enumerate(text):
        cp, name, cat = analyze_char(ch)
        char_freq[cp] = char_freq.get(cp, 0) + 1

        # Flag unusual characters
        is_devanagari = 0x0900 <= ord(ch) <= 0x097F
        is_ascii_printable = 0x20 <= ord(ch) <= 0x7E
        is_common = ch in '\n\r\t।॥' or is_devanagari or is_ascii_printable

        if not is_common:
            issues.append({
                "pos": i,
                "char": repr(ch),
                "codepoint": cp,
                "name": name,
                "category": cat,
                "context": text[max(0, i-10):i+10]
            })

    # Check for repetition patterns
    repetitions = []
    if len(text) > 20:
        # Check if any 10+ char substring repeats
        for length in [20, 15, 10]:
            for start in range(len(text) - length):
                substr = text[start:start + length]
                count = text.count(substr)
                if count >= 3:
                    repetitions.append({"pattern": repr(substr), "count": count, "length": length})
                    break
            if repetitions:
                break

    return issues, char_freq, repetitions


def smart_chunk(text, max_chars=400):
    marker_re = r"(?=\n[ \t]*(?:दफा|धारा|भाग|परिच्छेद|अनुसूची|Schedule|Part|Section|Article)\b)"
    parts = [p.strip() for p in re.split(marker_re, text) if p.strip()]
    if len(parts) <= 1:
        parts = [p.strip() for p in re.split(r"\n{2,}", text) if p.strip()]
    out, cur = [], ""
    for p in parts:
        if len(cur) + len(p) > max_chars and len(cur) > 0:
            out.append(cur.strip())
            cur = p
        else:
            cur += "\n" + p if cur else p
    if cur.strip():
        out.append(cur.strip())

    # Force-split oversized
    final = []
    for chunk in out:
        if len(chunk) > 600:
            sentences = re.split(r"(।)", chunk)
            sub, i = "", 0
            while i < len(sentences):
                piece = sentences[i]
                if i + 1 < len(sentences) and sentences[i + 1] == "।":
                    piece += "।"
                    i += 2
                else:
                    i += 1
                if len(sub) + len(piece) > 600 and sub:
                    final.append(sub.strip())
                    sub = piece
                else:
                    sub += piece
            if sub.strip():
                final.append(sub.strip())
        else:
            final.append(chunk)
    return final if final else [text]


if __name__ == "__main__":
    import sys

    filepath = "data/Nepal Ko Sanbidhan.docx"
    print(f"Analyzing: {filepath}\n")

    raw = extract_text_from_docx(filepath)
    cleaned = clean_ocr_text(raw)
    chunks = smart_chunk(cleaned, max_chars=400)

    print(f"Total chunks: {len(chunks)}\n")
    print("=" * 80)

    # Track overall stats
    total_issues = 0
    problematic_chunks = []

    for idx, chunk in enumerate(chunks):
        issues, char_freq, repetitions = analyze_text(chunk)

        # Only show detail for problematic chunks
        has_problems = len(issues) > 0 or len(repetitions) > 0

        if has_problems or len(chunk) < 20:
            problematic_chunks.append(idx)
            print(f"\n{'─' * 80}")
            print(f"CHUNK {idx + 1}/{len(chunks)}  |  {len(chunk)} chars")
            print(f"{'─' * 80}")

            # Show first 200 chars of the chunk
            preview = chunk[:200]
            print(f"Preview: {repr(preview)}")

            if issues:
                print(f"\n  ⚠️  UNUSUAL CHARACTERS ({len(issues)}):")
                for issue in issues[:10]:  # Show max 10
                    print(
                        f"    pos={issue['pos']}  char={issue['char']}  "
                        f"{issue['codepoint']}  {issue['name']}  "
                        f"cat={issue['category']}"
                    )
                    print(f"    context: ...{repr(issue['context'])}...")
                total_issues += len(issues)

            if repetitions:
                print(f"\n  🔁 REPETITION PATTERNS:")
                for rep in repetitions[:5]:
                    print(f"    {rep['pattern']} × {rep['count']}")

            if len(chunk) < 20:
                print(f"\n  📏 VERY SHORT CHUNK - Full content: {repr(chunk)}")
                print(f"    Char breakdown:")
                for ch in chunk:
                    cp, name, cat = analyze_char(ch)
                    print(f"      {repr(ch)}  {cp}  {name}  ({cat})")

    print(f"\n{'=' * 80}")
    print(f"SUMMARY")
    print(f"{'=' * 80}")
    print(f"Total chunks: {len(chunks)}")
    print(f"Problematic chunks: {len(problematic_chunks)}")
    print(f"Total unusual characters: {total_issues}")
    print(f"\nChunk size distribution:")

    size_buckets = {
        "1-10": 0, "11-50": 0, "51-100": 0, "101-200": 0,
        "201-400": 0, "401-600": 0, "600+": 0
    }
    for chunk in chunks:
        l = len(chunk)
        if l <= 10: size_buckets["1-10"] += 1
        elif l <= 50: size_buckets["11-50"] += 1
        elif l <= 100: size_buckets["51-100"] += 1
        elif l <= 200: size_buckets["101-200"] += 1
        elif l <= 400: size_buckets["201-400"] += 1
        elif l <= 600: size_buckets["401-600"] += 1
        else: size_buckets["600+"] += 1

    for bucket, count in size_buckets.items():
        bar = "█" * count
        print(f"  {bucket:>8}: {count:>4}  {bar}")
