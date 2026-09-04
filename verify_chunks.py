"""Quick check: How do the new chunks look after all fixes?"""
import re
from docx import Document

# Import the fixed functions
from ocr_correct import clean_ocr_text, smart_chunk, MAX_CHUNK_CHARS, HARD_CHAR_LIMIT, calc_max_tokens


def extract_text_from_docx(filepath):
    doc = Document(filepath)
    return "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])


if __name__ == "__main__":
    import sys
    from pathlib import Path

    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        data_dir = Path("data")
        docx_candidates = list(data_dir.glob("*.docx")) if data_dir.exists() else []
        if docx_candidates:
            filepath = str(docx_candidates[0])
            print(f"No file argument provided. Auto-selected: {filepath}\n")
        else:
            print("Usage: python verify_chunks.py [path_to_docx]")
            print("Notice: No .docx file specified and no .docx files found in 'data/' directory.")
            print("Please place a .docx file in 'data/' or specify the file path as an argument.")
            sys.exit(0)

    if not Path(filepath).exists():
        print(f"Error: File not found: {filepath}")
        sys.exit(1)

    raw = extract_text_from_docx(filepath)
    cleaned = clean_ocr_text(raw)
    chunks = smart_chunk(cleaned, max_chars=MAX_CHUNK_CHARS)

    print(f"Total chunks: {len(chunks)}")
    print(f"Config: MAX_CHUNK_CHARS={MAX_CHUNK_CHARS}, HARD_CHAR_LIMIT={HARD_CHAR_LIMIT}")

    # Stats
    sizes = [len(c) for c in chunks]
    oversized = [s for s in sizes if s > HARD_CHAR_LIMIT]
    print(f"\nSize stats: min={min(sizes)}, max={max(sizes)}, avg={sum(sizes)//len(sizes)}")
    print(f"Oversized (>{HARD_CHAR_LIMIT}): {len(oversized)}")

    # Token stats
    tokens = [calc_max_tokens(c) for c in chunks]
    print(f"\nmax_tokens stats: min={min(tokens)}, max={max(tokens)}, avg={sum(tokens)//len(tokens)}")

    # Distribution
    buckets = {"1-50": 0, "51-100": 0, "101-200": 0, "201-300": 0, "301-400": 0, "401-500": 0, "500+": 0}
    for s in sizes:
        if s <= 50: buckets["1-50"] += 1
        elif s <= 100: buckets["51-100"] += 1
        elif s <= 200: buckets["101-200"] += 1
        elif s <= 300: buckets["201-300"] += 1
        elif s <= 400: buckets["301-400"] += 1
        elif s <= 500: buckets["401-500"] += 1
        else: buckets["500+"] += 1

    print("\nChunk size distribution:")
    for bucket, count in buckets.items():
        bar = "#" * min(count, 80)
        print(f"  {bucket:>8}: {count:>4}  {bar}")

    # Check for ZWNJ in cleaned text
    zwnj_count = cleaned.count("\u200c")
    print(f"\nZWNJ chars remaining after clean: {zwnj_count}")

    # Show a few sample chunks with their max_tokens
    print("\n--- Sample chunks ---")
    for i in [0, 1, len(chunks)//2, len(chunks)-1]:
        c = chunks[i]
        mt = calc_max_tokens(c)
        print(f"\nChunk {i+1}: {len(c)} chars -> max_tokens={mt}")
        print(f"  Preview: {repr(c[:120])}...")
