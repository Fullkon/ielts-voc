# -*- coding: utf-8 -*-
"""
IELTS Test Bank Generator — LLM-powered question authoring.

Generates 4 question types per day:
  cn2en       — 汉译英: purely from Excel (Chinese def -> English word)
  en2cn       — 英译汉: purely from Excel (English word -> Chinese def)
  collocation — 短语搭配翻译: LLM picks a collocation, blanks target word, adds Chinese hint
  sentence    — 句子词汇填空: LLM writes a sentence, blanks target word, adds Chinese hint

Output:  test_bank_llm/day_{day:02d}.json  (one file per day, 28 files total)
Checkpoint:  test_bank_llm/checkpoint.json  (tracks which days are done)

LLM: Volcano API, doubao-seed-2-1-pro-260628, batch=25 words, JSON-only output.
"""

import json
import time
import sys
import re
from pathlib import Path
import pandas as pd
import requests

# ═════════════════════════════════════════════════════════════════════════════
try:
    from config import ARK_KEY, ARK_URL, MODEL
except ImportError:
    ARK_KEY = "your-ark-api-key-here"
    ARK_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
    MODEL   = "doubao-seed-2-1-pro-260628"

XLSX_FILE   = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'
BANK_DIR    = Path(__file__).parent / 'test_bank_llm'
CHECKPOINT  = BANK_DIR / 'checkpoint.json'
TOTAL_DAYS  = 28
BATCH_SIZE  = 25
MAX_RETRIES = 3
RETRY_DELAY = 15
TIMEOUT     = (60, 600)

# ═════════════════════════════════════════════════════════════════════════════
# LEARNING PLAN (same as app.py)
# ═════════════════════════════════════════════════════════════════════════════

def make_plan(total: int) -> dict[int, list[int]]:
    base = total // TOTAL_DAYS
    rem  = total % TOTAL_DAYS
    plan = {}
    start = 0
    for day in range(1, TOTAL_DAYS + 1):
        count = base + (1 if day <= rem else 0)
        plan[day] = list(range(start, start + count))
        start += count
    return plan

# ═════════════════════════════════════════════════════════════════════════════
# DIRECT questions (cn2en, en2cn) — no LLM needed
# ═════════════════════════════════════════════════════════════════════════════

def make_cn2en_questions(df: pd.DataFrame, word_ids: list[int]) -> list[dict]:
    qs = []
    for wid in word_ids:
        row = df.iloc[wid]
        word = str(row['word']).strip()
        cn   = str(row['chinese_def']).strip()
        hint = f"{word[0].upper()}..." if len(word) > 3 else f"{len(word)} letters"
        qs.append({
            "word_id": wid,
            "question": cn,
            "answer": word.lower(),
            "hint": hint,
        })
    return qs

def make_en2cn_questions(df: pd.DataFrame, word_ids: list[int]) -> list[dict]:
    qs = []
    for wid in word_ids:
        row = df.iloc[wid]
        word = str(row['word']).strip()
        cn   = str(row['chinese_def']).strip()
        en_def = str(row['english_def']).split('|')[0].strip()[:120]
        qs.append({
            "word_id": wid,
            "question": word,
            "answer": cn.lower(),
            "hint": en_def if en_def else "Please translate into Chinese",
        })
    return qs

# ═════════════════════════════════════════════════════════════════════════════
# LLM PROMPT
# ═════════════════════════════════════════════════════════════════════════════

def make_llm_prompt(words_info: list[dict]) -> str:
    """Build a prompt listing words with their collocations and Chinese def."""
    items = []
    for wi in words_info:
        colls = wi.get("collocations", "")
        cn    = wi.get("chinese_def", "")
        items.append(
            f'  {{"word":"{wi["word"]}","collocations":{json.dumps(colls)},'
            f'"chinese":"{cn}"}}'
        )
    return (
        "You are an IELTS test question author.\n\n"
        "For each word below, create TWO questions:\n\n"
        "1) collocation — Pick ONE collocation phrase, replace the target word "
        "with _____ and add a Chinese hint in parentheses.\n"
        '   Example: "_____ chamber (岩浆房)" → {'
        '"word":"magma","coll_question":"_____ chamber (岩浆房)","coll_answer":"magma"}\n\n'
        "2) sentence — Write ONE natural English sentence using the word, "
        "replace the target word with _____ and append the Chinese meaning in parentheses.\n"
        '   Example: {'
        '"word":"magma","sent_question":"The _____ beneath the volcano is moving. (岩浆)","sent_answer":"magma"}\n\n'
        "Words:\n" + "\n".join(items) + "\n\n"
        "Output ONLY a JSON array, one object per word, each with: "
        'word, coll_question, coll_answer, sent_question, sent_answer. '
        "No other text."
    )

# ═════════════════════════════════════════════════════════════════════════════
# API CALL
# ═════════════════════════════════════════════════════════════════════════════

def call_llm(words_info: list[dict]) -> list[dict]:
    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": make_llm_prompt(words_info)}],
        "max_tokens": 8192,
        "temperature": 0.3,
        "stream": False,
    }
    h = {"Authorization": f"Bearer {ARK_KEY}", "Content-Type": "application/json"}

    for attempt in range(1, MAX_RETRIES + 1):
        t0 = time.time()
        try:
            r = requests.post(ARK_URL, json=payload, headers=h, timeout=TIMEOUT)
            elapsed = time.time() - t0
            print(f"    HTTP {r.status_code}  {elapsed:.0f}s", end="", flush=True)
            if r.status_code == 200:
                raw = r.json()["choices"][0]["message"]["content"]
                items = _extract_json_array(raw, words_info)
                print(f"  parsed {len(items)}/{len(words_info)} words", flush=True)
                return items
            if r.status_code == 429:
                wait = RETRY_DELAY * 3
                print(f"  rate-limited, wait {wait}s", flush=True)
                time.sleep(wait)
            elif r.status_code >= 500:
                print(f"  server err, retry {attempt}/{MAX_RETRIES}", flush=True)
                time.sleep(RETRY_DELAY)
            else:
                print(f"\n    Body: {r.text[:200]}", flush=True)
                time.sleep(RETRY_DELAY)
        except Exception as e:
            print(f"  {e}", flush=True)
            time.sleep(RETRY_DELAY)
    print("  [FAILED]", flush=True)
    return []

def _extract_json_array(text: str, words_info: list[dict]) -> list[dict]:
    text = re.sub(r'```(?:json)?\s*', '', text)
    lo = text.find('[')
    hi = text.rfind(']')
    if lo < 0 or hi <= lo:
        print("  no array found", flush=True)
        return []
    text = text[lo:hi+1]
    # Remove control characters (ASCII 0-31) except \n, \r, \t
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    try:
        raw = json.loads(text)
    except Exception as e:
        print(f"  JSON error: {e}", flush=True)
        return []

    # Build lookup
    lookup = {}
    for wi in words_info:
        lookup[wi["word"].lower()] = wi["word_idx"]

    results = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        wkey = (item.get("word") or "").strip().lower()
        wid = lookup.get(wkey)
        if wid is None:
            for w_lower, idx in lookup.items():
                if wkey in w_lower or w_lower in wkey:
                    wid = idx
                    break
        if wid is None:
            continue
        results.append({
            "word_id": wid,
            "coll_question": (item.get("coll_question") or "").strip(),
            "coll_answer":   (item.get("coll_answer") or "").strip(),
            "sent_question": (item.get("sent_question") or "").strip(),
            "sent_answer":   (item.get("sent_answer") or "").strip(),
        })
    return results

# ═════════════════════════════════════════════════════════════════════════════
# DAY GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

def generate_day(df: pd.DataFrame, day: int, word_ids: list[int]):
    """Generate all 4 test types for one day, save to JSON."""
    print(f"\n{'='*60}")
    print(f"Day {day}: {len(word_ids)} words")
    print(f"{'='*60}")

    # Step 1: Direct questions (cn2en, en2cn)
    cn2en = make_cn2en_questions(df, word_ids)
    en2cn = make_en2cn_questions(df, word_ids)
    print(f"  cn2en: {len(cn2en)} items (direct)")
    print(f"  en2cn: {len(en2cn)} items (direct)")

    # Step 2: LLM-generated (collocation + sentence)
    # Build word info list
    all_word_info = []
    for wid in word_ids:
        row = df.iloc[wid]
        colls = parse_collocations(str(row['collocations']))
        all_word_info.append({
            "word_idx": wid,
            "word": str(row['word']).strip(),
            "collocations": colls,
            "chinese_def": str(row['chinese_def']).strip(),
        })

    # Batch
    coll_questions = []
    sent_questions = []
    total_llm = 0
    t_llm_start = time.time()

    for bi in range(0, len(all_word_info), BATCH_SIZE):
        batch = all_word_info[bi:bi + BATCH_SIZE]
        bn = bi // BATCH_SIZE + 1
        total_bn = (len(all_word_info) + BATCH_SIZE - 1) // BATCH_SIZE
        print(f"\n  [{bn}/{total_bn}] batch {len(batch)} words:", end=" ", flush=True)
        print(", ".join(f"{wi['word']}" for wi in batch[:4]) + ("..." if len(batch) > 4 else ""), flush=True)

        items = call_llm(batch)
        for item in items:
            coll_questions.append({
                "word_id": item["word_id"],
                "question": item["coll_question"],
                "answer":   item["coll_answer"].lower(),
                "hint":     "请根据中文提示填写搭配中缺失的词",
            })
            sent_questions.append({
                "word_id": item["word_id"],
                "question": item["sent_question"],
                "answer":   item["sent_answer"].lower(),
                "hint":     "请根据上下文和中文提示填空",
            })
        total_llm += len(items)

        time.sleep(5)  # rate limit

    print(f"\n  collocation (LLM): {len(coll_questions)} / {len(word_ids)}")
    print(f"  sentence (LLM):   {len(sent_questions)} / {len(word_ids)}")
    print(f"  LLM time: {(time.time()-t_llm_start)/60:.0f}m")

    # Save
    BANK_DIR.mkdir(parents=True, exist_ok=True)
    day_file = BANK_DIR / f"day_{day:02d}.json"
    data = {
        "day": day,
        "word_count": len(word_ids),
        "questions": {
            "cn2en":       cn2en,
            "en2cn":       en2cn,
            "collocation": coll_questions,
            "sentence":    sent_questions,
        }
    }
    day_file.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"  Saved -> {day_file.name}\n")

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def parse_collocations(s: str) -> list:
    if not s or s in ('None', 'nan', ''):
        return []
    return [p.strip() for p in re.split(r'[,|，]', s) if p.strip()]

# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("IELTS Test Bank Generator (LLM-powered)")
    print(f"Model: {MODEL}   Batch: {BATCH_SIZE} words")
    print("=" * 60)

    df = pd.read_excel(XLSX_FILE)
    df.columns = ['index_no', 'word', 'pronunciation', 'english_def',
                  'collocations', 'sentence', 'root_words', 'related_words', 'chinese_def']
    df['word'] = df['word'].astype(str).str.strip()
    df['chinese_def'] = df['chinese_def'].astype(str).str.strip()
    df['english_def'] = df['english_def'].fillna('').astype(str)

    total = len(df)
    plan = make_plan(total)
    print(f"Words: {total}  Days: {TOTAL_DAYS}")

    # Checkpoint
    BANK_DIR.mkdir(parents=True, exist_ok=True)
    done_days = set()
    if CHECKPOINT.exists():
        try:
            done_days = set(json.loads(CHECKPOINT.read_text(encoding='utf-8')).get("done", []))
        except:
            pass

    # Check existing day files
    for day in range(1, TOTAL_DAYS + 1):
        f = BANK_DIR / f"day_{day:02d}.json"
        if f.exists():
            try:
                d = json.loads(f.read_text(encoding='utf-8'))
                needed = len(plan[day])
                have = max(len(d.get("questions", {}).get(t, [])) for t in ["cn2en","en2cn","collocation","sentence"])
                if have >= needed:
                    done_days.add(day)
            except:
                pass

    pending = [d for d in range(1, TOTAL_DAYS + 1) if d not in done_days]
    print(f"Done: {len(done_days)} days  |  Pending: {len(pending)} days")

    if not pending:
        print("[DONE] All days complete!")
        return

    # Estimate
    est_calls = sum(
        (len(plan[d]) + BATCH_SIZE - 1) // BATCH_SIZE for d in pending
    )
    est_h = est_calls * 320 / 3600
    print(f"Est. LLM calls: {est_calls}  |  Est. time: ~{est_h:.1f}h\n")

    t_total = time.time()
    for di, day in enumerate(pending):
        word_ids = plan[day]
        generate_day(df, day, word_ids)

        done_days.add(day)
        CHECKPOINT.write_text(
            json.dumps({"done": sorted(done_days)}, ensure_ascii=False),
            encoding='utf-8')

        elapsed_m = (time.time() - t_total) / 60
        eta_m = (elapsed_m / (di + 1)) * (len(pending) - di - 1) if di < len(pending) else 0
        print(f"[TIME] elapsed {elapsed_m:.0f}m  |  eta {eta_m:.0f}m remaining")

    total_m = (time.time() - t_total) / 60
    print(f"\n[DONE] All {TOTAL_DAYS} days generated in {total_m:.0f}m ({total_m/60:.1f}h)")

if __name__ == '__main__':
    main()
