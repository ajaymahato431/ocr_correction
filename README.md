# OCR Correction

A Python pipeline that corrects OCR errors in Nepali legal `.docx` documents using an OpenAI-compatible LLM API. It cleans scan artifacts, chunks documents along legal structure boundaries (दफा/धारा/भाग), sends each chunk to the model for correction, validates the output against hallucination/truncation patterns, and reassembles a formatted `.docx` with proper headings and indentation.

## Features

- **Pre-cleaning**: strips zero-width characters, smart quotes, watermark/URL noise, and other scan artifacts before chunking.
- **Structure-aware chunking**: splits text on legal markers (दफा, धारा, भाग, परिच्छेद, अनुसूची, Schedule, Part, Section, Article) with a hard character ceiling and sentence/newline fallback splitting.
- **Multi-threaded correction**: processes chunks in parallel via `ThreadPoolExecutor`, with a retry pass for chunks that fail.
- **Output validation**: rejects and recursively re-splits chunks that come back truncated, empty, oversized, or repetitive (a known LLM hallucination pattern).
- **Culprit sanitization**: detects individual characters/substrings that blow up token counts (bad OCR glyphs) and strips them before sending to the API, logging what was dropped.
- **DOCX reconstruction**: rebuilds a `.docx` with headings for भाग/अनुसूची, bold titles for दफा, and indentation for उपदफा/खण्ड.
- **Selectable API provider**: switch between configured OpenAI-compatible providers (DeepSeek, or a custom one) via `.env` or interactive prompt.

## Requirements

- Python 3.9+
- Dependencies in [requirements.txt](requirements.txt): `openai`, `python-docx`, `python-dotenv`, `tiktoken`

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` and set your API key:

```
DEEPSEEK_API_KEY=your_deepseek_api_key_here
DEFAULT_API_PROVIDER=deepseek
```

## Usage

1. Place source `.docx` files in a `data/` folder in the project root.
2. Run the script:

```bash
python ocr_correct.py
```

3. Choose an API provider when prompted (or press Enter to use `DEFAULT_API_PROVIDER` from `.env`).
4. Corrected files are written to `output/` as `<original_name>_corrected.docx`.

Errors and skipped/unrecoverable text segments are logged to `error_logs.txt` and `skipped_culprits.txt` in the project root.

## Configuration

Key constants at the top of [ocr_correct.py](ocr_correct.py):

| Constant | Purpose |
|---|---|
| `MAX_CHUNK_CHARS` | Target chunk size (default 1500 chars) |
| `HARD_CHAR_LIMIT` | Absolute per-chunk ceiling (default 2000 chars) |
| `MAX_OUTPUT_TOKENS` | Max tokens requested per API call (default 8192) |
| `MAX_WORKERS` | Parallel API request threads (default 5) |
| `API_PROVIDERS` | Provider name → API key env var, base URL, and model |

To add another OpenAI-compatible provider, add an entry to `API_PROVIDERS` in [ocr_correct.py](ocr_correct.py) with its own `api_key_env`, `base_url`, and `model`, plus a matching key in `.env`.

## Helper scripts

- [diagnose_chunks.py](diagnose_chunks.py) — dumps chunk-by-chunk analysis (unusual characters, repetition patterns, size distribution) for a hardcoded input file, useful for debugging bad OCR input before a full run.
- [verify_chunks.py](verify_chunks.py) — quick sanity check of chunk sizing and token estimates for a hardcoded input file after tuning the chunking constants.

Both scripts currently point at `data/Nepal Ko Sanbidhan.docx`; edit the `filepath` variable in each to point at a different file.

## How it works

1. **Extract**: paragraph text is pulled from the source `.docx`.
2. **Clean**: `clean_ocr_text` normalizes characters and drops noise lines (page numbers, URLs, short non-Devanagari fragments).
3. **Chunk**: `smart_chunk` splits on legal structure markers, merges small pieces up to `MAX_CHUNK_CHARS`, and force-splits anything over `HARD_CHAR_LIMIT`.
4. **Correct**: each chunk is sent to the LLM with a system prompt tuned for Nepali legal OCR correction (preserves numbering hierarchy, punctuation, and wording; forbids paraphrasing/translation).
5. **Validate**: `validate_output` checks output length ratio and repetition; invalid output is recursively bisected and retried until a clean segment is isolated or the culprit is skipped and logged.
6. **Reassemble**: corrected chunks are joined, paragraph breaks are fixed, दफा/उपदफा/खण्ड boundaries are reformatted onto their own lines, and the result is rebuilt into a formatted `.docx`.
