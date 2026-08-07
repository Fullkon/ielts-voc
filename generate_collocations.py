# -*- coding: utf-8 -*-
"""
IELTS Collocation Phrase Generator — Volcano API (doubao-seed-2-1-pro-260628)

Generates proper collocation PHRASES (adj+noun, noun+noun, verb+noun, etc.)
for each IELTS vocabulary word, writing results to Excel with full checkpoint/resume.

Key design decisions:
- Batch = 50 words per call (maximized because model has fixed ~280s reasoning overhead)
- Non-streaming mode for clean JSON parsing
- Structured JSON output enforced via concise prompt
- Checkpoint saved after every batch; Excel updated every 5 batches
- 5 s delay between batches (basic rate limiting)
- 3 retry attempts with exponential backoff on failure
"""

import json
import time
import sys
import re
from pathlib import Path
import pandas as pd
import requests

# ═════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════

try:
    from config import ARK_KEY, ARK_URL, MODEL
except ImportError:
    ARK_KEY = "your-ark-api-key-here"
    ARK_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
    MODEL   = "doubao-seed-2-1-pro-260628"

XLSX_FILE   = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'
CHECKPOINT  = Path(__file__).parent / 'collocations_phrases_checkpoint.json'
OUTPUT_JSON = Path(__file__).parent / 'collocations_phrases.json'

BATCH_SIZE  = 50          # words per API call (model fixed overhead ~280s, so maximize)
MAX_RETRIES = 3           # attempts per batch
RETRY_DELAY = 15          # seconds between retries
BATCH_DELAY = 5           # seconds between successive batches
TIMEOUT     = (60, 600)   # (connect, read) timeouts

# ═════════════════════════════════════════════════════════════════════════════
# PROMPT
# ═════════════════════════════════════════════════════════════════════════════

def make_prompt(words: list[str]) -> str:
    words_json = json.dumps(words, ensure_ascii=False)
    return (
        "For each word below, give 3-5 natural English collocation PHRASES "
        "(adj+noun, noun+noun, verb+noun, adv+verb). "
        "Each phrase must contain at least TWO content words. "
        "Output ONLY a JSON object — no markdown, no extra text.\n\n"
        "Words: " + words_json + "\n\n"
        'Format: {"word":["phrase1","phrase2","phrase3"]}'
    )


# ═════════════════════════════════════════════════════════════════════════════
# API CALL
# ═════════════════════════════════════════════════════════════════════════════

def call_api(words: list[str], batch_num: int) -> dict[str, list[str]]:
    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": make_prompt(words)}],
        "max_tokens": 8192,
        "temperature": 0.3,
        "stream": False,
    }
    headers = {
        "Authorization": f"Bearer {ARK_KEY}",
        "Content-Type": "application/json",
    }

    for attempt in range(1, MAX_RETRIES + 1):
        t0 = time.time()
        try:
            resp = requests.post(
                ARK_URL, json=payload, headers=headers, timeout=TIMEOUT)
            elapsed = time.time() - t0
            print(f"    HTTP {resp.status_code}  {elapsed:.0f}s",
                  end="", flush=True)

            if resp.status_code == 200:
                data = resp.json()
                raw = data["choices"][0]["message"]["content"].strip()
                parsed = _extract_json(raw, words)
                print(f"  parsed {len(parsed)}/{len(words)} words", flush=True)
                return parsed

            if resp.status_code == 429:
                wait = RETRY_DELAY * 3
                print(f"  rate-limited, waiting {wait}s", flush=True)
                time.sleep(wait)
            elif resp.status_code >= 500:
                print(f"  server error, retry {attempt}/{MAX_RETRIES}",
                      flush=True)
                time.sleep(RETRY_DELAY)
            else:
                snippet = resp.text[:200].replace("\n", " ")
                print(f"\n    Body: {snippet}", flush=True)
                time.sleep(RETRY_DELAY)

        except requests.exceptions.Timeout:
            print(f"  timeout, retry {attempt}/{MAX_RETRIES}", flush=True)
            time.sleep(RETRY_DELAY * 2)
        except Exception as e:
            print(f"  error: {e}", flush=True)
            time.sleep(RETRY_DELAY)

    print("    FAILED after all retries", flush=True)
    return {}


# ═════════════════════════════════════════════════════════════════════════════
# RESPONSE PARSER
# ═════════════════════════════════════════════════════════════════════════════

def _extract_json(text: str, words: list[str]) -> dict[str, list[str]]:
    # Strip markdown code fences
    text = re.sub(r'```(?:json)?\s*', '', text)
    text = re.sub(r'```$', '', text)

    # Locate outermost { … }
    lo = text.find('{')
    hi = text.rfind('}')
    if lo >= 0 and hi > lo:
        text = text[lo:hi + 1]

    # Parse
    try:
        raw = json.loads(text)
    except json.JSONDecodeError as e:
        print(f"\n    [WARN] JSON decode: {e}", flush=True)
        print(f"    Raw (first 400 chars): {text[:400]}", flush=True)
        return {}

    # Build case-insensitive lookup
    lookup: dict[str, str] = {}
    for w in words:
        lookup[w.lower()] = w
        # also index by simplified key
        simple = re.sub(r'[\s\-\(\)（）/]', '', w.lower())
        if simple and simple != w.lower():
            lookup[simple] = w

    result: dict[str, list[str]] = {}
    for key, val in raw.items():
        key_stripped = key.strip()
        key_lower = key_stripped.lower()
        # exact match
        orig = lookup.get(key_lower)
        # fuzzy match
        if orig is None:
            simple = re.sub(r'[\s\-\(\)（）/]', '', key_lower)
            orig = lookup.get(simple)
        # containment match
        if orig is None:
            for w in words:
                if key_lower in w.lower() or w.lower() in key_lower:
                    orig = w
                    break
        if orig is None:
            orig = key_stripped   # keep as-is

        if isinstance(val, list):
            phrases = []
            for v in val:
                if isinstance(v, str) and v.strip():
                    phrases.append(v.strip())
            result[orig] = list(dict.fromkeys(phrases))   # dedupe, keep order
        elif isinstance(val, str) and val.strip():
            result[orig] = [val.strip()]

    return result


# ═════════════════════════════════════════════════════════════════════════════
# EXCEL HELPER
# ═════════════════════════════════════════════════════════════════════════════

def _write_excel(progress: dict, df: pd.DataFrame, path: Path, total: int):
    updated = 0
    for i in range(total):
        w = str(df.at[i, 'word']).strip()
        if w in progress:
            val = progress[w]
            if isinstance(val, list):
                df.at[i, 'collocations'] = ", ".join(val)
            else:
                df.at[i, 'collocations'] = str(val)
            updated += 1

    out = pd.DataFrame({
        '序号':       df['index_no'],
        '英文单词':   df['word'],
        '英文释义':   df['english_def'],
        '汉语释义':   df['chinese_def'],
        '常见搭配':   df['collocations'],
        '例句':       df['sentence'],
        '同根词':     df['root_words'],
        '相关同类词': df['related_words'],
        '备注':       df['notes'],
    })
    out.to_excel(path, index=False)
    print(f"    [SAVE] Excel saved ({updated}/{total} words have collocations)",
          flush=True)


# ═════════════════════════════════════════════════════════════════════════════
# PROGRESS DISPLAY
# ═════════════════════════════════════════════════════════════════════════════

def _show_sample(progress: dict, n: int = 5):
    items = list(progress.items())
    if not items:
        return
    print("    Sample entries:")
    for w, c in items[-n:]:
        preview = ", ".join(c[:4]) if isinstance(c, list) else str(c)[:80]
        print(f"      {w} -> {preview}")


# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print(f"IELTS Collocation Phrase Generator")
    print(f"Model: {MODEL}   Batch: {BATCH_SIZE} words   JSON output")
    print("=" * 60)

    # ── Load Excel ──
    if not XLSX_FILE.exists():
        print(f"[ERROR] Excel not found: {XLSX_FILE}")
        sys.exit(1)

    df = pd.read_excel(XLSX_FILE)
    df.columns = ['index_no', 'word', 'english_def', 'chinese_def',
                  'collocations', 'sentence', 'root_words', 'related_words', 'notes']
    df['word'] = df['word'].astype(str).str.strip()
    df['collocations'] = df['collocations'].fillna('').astype(str)
    total = len(df)
    print(f"Words in Excel: {total}")

    # ── Load checkpoint ──
    progress: dict[str, list[str]] = {}
    if CHECKPOINT.exists():
        try:
            progress = json.loads(CHECKPOINT.read_text(encoding='utf-8'))
            if not isinstance(progress, dict):
                progress = {}
        except Exception:
            print("[WARN] Corrupt checkpoint — starting fresh.")
            progress = {}

    # ── Validate checkpoint format (must be list of phrases, not single words) ──
    if progress:
        sample_word = next(iter(progress))
        sample_val = progress[sample_word]
        is_phrase = (
            isinstance(sample_val, list) and len(sample_val) > 0
            and isinstance(sample_val[0], str)
            and " " in sample_val[0]
        )
        if not is_phrase:
            print("[WARN] Old-format checkpoint (single words, not phrases). "
                  "Starting fresh.")
            progress = {}

    done = len(progress)
    print(f"Already processed: {done}  |  Remaining: {total - done}")

    if done == total:
        print("[OK] All done! Writing final Excel ...")
        _write_excel(progress, df, XLSX_FILE, total)
        return

    # ── Build batches ──
    remaining: list[str] = []
    for i in range(total):
        w = str(df.at[i, 'word']).strip()
        if w and w not in progress:
            remaining.append(w)

    batches: list[list[str]] = []
    for i in range(0, len(remaining), BATCH_SIZE):
        batches.append(remaining[i:i + BATCH_SIZE])

    NB = len(batches)
    est_h = NB * 320 / 3600  # ~320s per batch (model reasoning overhead is the bottleneck)
    print(f"Batches to process: {NB}   Est. time: ~{est_h:.1f} h\n")

    # ── Loop ──
    t_total = time.time()
    for bi, batch in enumerate(batches):
        bn = bi + 1
        print(f"[{bn}/{NB}]  {len(batch)} words: "
              f"{batch[0]} ... {batch[-1]}", flush=True)

        result = call_api(batch, bn)

        if result:
            progress.update(result)
            pct = 100 * len(progress) / total
            print(f"    [OK] +{len(result)}  |  "
                  f"total {len(progress)}/{total} ({pct:.1f}%)", flush=True)
        else:
            print("    [ERROR] batch failed — will retry on next run", flush=True)

        # ── Save checkpoint every batch ──
        CHECKPOINT.write_text(
            json.dumps(progress, ensure_ascii=False, indent=2),
            encoding='utf-8')

        # ── Update Excel every 5 batches ──
        if bn % 5 == 0 or bn == NB:
            _write_excel(progress, df, XLSX_FILE, total)

        # ── ETA ──
        elapsed_m = (time.time() - t_total) / 60
        eta_m = (elapsed_m / bn) * (NB - bn) if bn > 0 else 0
        print(f"    [TIME] elapsed {elapsed_m:.0f}m  |  "
              f"eta {eta_m:.0f}m remaining\n", flush=True)

        # ── Rate limit ──
        if bn < NB:
            time.sleep(BATCH_DELAY)

    # ── Final ──
    total_m = (time.time() - t_total) / 60
    print(f"\n{'=' * 60}")
    print(f"[DONE] COMPLETE  ({total_m:.0f}m / {total_m/60:.1f}h)")
    print(f"   Words with collocation phrases: {len(progress)} / {total}")
    _show_sample(progress)
    print()

    # Save JSON output
    OUTPUT_JSON.write_text(
        json.dumps(progress, ensure_ascii=False, indent=2), encoding='utf-8')

    # Final Excel
    _write_excel(progress, df, XLSX_FILE, total)
    print("[OK] All files saved.")


if __name__ == '__main__':
    main()
