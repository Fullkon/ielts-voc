# -*- coding: utf-8 -*-
"""
Regenerate IELTS word collocations using Volcano LLM — sequential, reliable.
Strict: content-word collocations only, NO function words.
Checkpoint-resumable. Runs each batch sequentially with streaming.
"""
import json, time, sys
from pathlib import Path
import pandas as pd
import requests

try:
    from config import ARK_KEY, ARK_URL, MODEL
except ImportError:
    ARK_KEY = "your-ark-api-key-here"
    ARK_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
    MODEL   = "doubao-seed-2-1-pro-260628"

XLSX_FILE = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'
CHECKPOINT = Path(__file__).parent / 'collocations_checkpoint.json'
OUTPUT_JSON = Path(__file__).parent / 'collocations_new.json'

BATCH_SIZE = 100
TIMEOUT = 1800  # 30 min per call (actual ~6 min)

def make_prompt(words):
    wl = "\n".join(f"- {w}" for w in words)
    return f"""{wl}

For each word, output one line: word=colloc1,colloc2,colloc3
Each collocation MUST be a two-word content pair: noun+noun (e.g., climate change), adj+noun (e.g., renewable energy), or verb+noun (e.g., conduct research).
Phrasal verbs or adverb modifiers are acceptable for verbs only.
NO function words: no prepositions, articles, or infinitive "to".
Prefer academic/terminological/IELTS-relevant combinations.
Output {len(words)} lines exactly. No explanations."""

def call_api(words, batch_num):
    """Call streaming API. Returns {word: colloc_string}."""
    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": make_prompt(words)}],
        "max_tokens": 4096,
        "stream": True
    }
    h = {"Authorization": f"Bearer {ARK_KEY}", "Content-Type": "application/json"}
    
    start = time.time()
    try:
        r = requests.post(ARK_URL, json=payload, headers=h,
                         timeout=(60, TIMEOUT), stream=True)
        if r.status_code != 200:
            print(f"    HTTP {r.status_code}: {r.text[:100]}")
            return {}
        
        full = ""
        for line in r.iter_lines(decode_unicode=True):
            if not line or not line.startswith("data: "):
                continue
            d = line[6:]
            if d.strip() == "[DONE]":
                break
            try:
                chunk = json.loads(d)
                full += chunk['choices'][0].get('delta', {}).get('content', '')
            except:
                pass
        
        elapsed = time.time() - start
        result = {}
        for line in full.strip().split('\n'):
            line = line.strip()
            if '=' in line:
                w, c = line.split('=', 1)
                w, c = w.strip(), c.strip()
                for ow in words:
                    if ow.lower() == w.lower() or w.lower() in ow.lower() or ow.lower() in w.lower():
                        result[ow] = c
                        break
        
        print(f"    {elapsed:.0f}s, {len(result)}/{len(words)} words parsed")
        return result
    
    except Exception as e:
        print(f"    Error ({time.time()-start:.0f}s): {e}")
        return {}

def main():
    print(f"Collocation Regenerator — {MODEL}")
    print(f"Batch: {BATCH_SIZE} words, Sequential streaming")
    print("=" * 50)
    
    df = pd.read_excel(XLSX_FILE)
    df.columns = ['index_no', 'word', 'pronunciation', 'english_def',
                   'collocations', 'sentence', 'root_words', 'related_words', 'chinese_def']
    df['word'] = df['word'].astype(str).str.strip()
    total = len(df)
    
    progress = {}
    if CHECKPOINT.exists():
        progress = json.load(open(CHECKPOINT, encoding='utf-8'))
    
    done = len(progress)
    print(f"Total: {total}, Already done: {done}")
    
    # Build remaining batches
    batches = []
    batch = []
    for i in range(total):
        w = str(df.iloc[i]['word']).strip()
        if w not in progress:
            batch.append(w)
            if len(batch) >= BATCH_SIZE:
                batches.append(batch)
                batch = []
    if batch:
        batches.append(batch)
    
    if not batches:
        print("All done! Updating Excel...")
    else:
        nb = len(batches)
        est = nb * (TIMEOUT * 0.5) / 3600
        print(f"Remaining batches: {nb} | Est: upto {est:.1f}h")
        print()
        
        total_start = time.time()
        for bi, b in enumerate(batches):
            bnum = bi + 1
            print(f"[{bnum}/{nb}] {len(b)} words: {b[0]}...{b[-1]}", flush=True)
            
            result = call_api(b, bnum)
            
            if result:
                progress.update(result)
                hits = len(result)
                print(f"    OK: {hits} words | Progress: {len(progress)}/{total} ({100*len(progress)/total:.1f}%)")
            else:
                print(f"    FAILED")
            
            # Save checkpoint
            json.dump(progress, open(CHECKPOINT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
            
            elapsed = (time.time() - total_start) / 60
            remaining = nb - bnum
            eta = (elapsed / bnum) * remaining if bnum > 0 else 0
            print(f"    Elapsed: {elapsed:.0f}m | ETA: {eta:.0f}m remaining")
            print()
            
            # Rate limiting
            if bnum < nb:
                time.sleep(5)
        
        total_time = (time.time() - total_start) / 60
        print(f"\nTotal time: {total_time:.0f}m ({total_time/60:.1f}h)")
    
    # Save output
    json.dump(progress, open(OUTPUT_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    
    # Update Excel
    print("Updating Excel...")
    updated = 0
    for i in range(total):
        w = str(df.iloc[i]['word']).strip()
        if w in progress:
            df.at[i, 'collocations'] = progress[w]
            updated += 1
    
    df_out = pd.DataFrame({
        '序号': df['index_no'], '英文单词': df['word'],
        '英文发音': df['pronunciation'],
        '英文释义': df['english_def'], '常见搭配': df['collocations'],
        '例句': df['sentence'],
        '同根词': df['root_words'], '相近词': df['related_words'],
        '汉语释义': df['chinese_def']
    })
    df_out.to_excel(XLSX_FILE, index=False)
    print(f"Excel updated: {updated}/{total} words. Done!")

if __name__ == '__main__':
    main()
