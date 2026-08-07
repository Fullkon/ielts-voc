# -*- coding: utf-8 -*-
"""Test mega-batch: 200 words in one API call."""
import json, requests, time, pandas as pd
from pathlib import Path

try:
    from config import ARK_KEY, ARK_URL
except ImportError:
    ARK_KEY = "your-ark-api-key-here"
    ARK_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

# Load first 200 words
xlsx = Path(__file__).parent / '雅思英文词汇表（完整版）.xlsx'
df = pd.read_excel(xlsx)
df.columns = ['index_no', 'word', 'english_def', 'chinese_def',
               'collocations', 'sentence', 'root_words', 'related_words', 'notes']
df['word'] = df['word'].astype(str).str.strip()
words = [str(df.iloc[i]['word']).strip() for i in range(100)]

wl = "\n".join(f"- {w}" for w in words)
prompt = f"""{wl}

For each word above, one line: word=colloc1,colloc2,colloc3
RULES: content-word collocations ONLY. Noun+noun, adj+noun, verb+noun, adv+verb.
NO function words (prepositions, articles, "to"). Academic/terminological.
Output ONLY {len(words)} lines. No explanations, no numbering."""

print(f"Mega-batch test: {len(words)} words")
print(f"Prompt length: {len(prompt)} chars")
start = time.time()

payload = {
    "model": "doubao-seed-2-1-pro-260628",
    "messages": [{"role": "user", "content": prompt}],
    "max_tokens": 4096,
    "stream": True
}

try:
    r = requests.post(ARK_URL, json=payload,
        headers={"Authorization": f"Bearer {ARK_KEY}", "Content-Type": "application/json"},
        timeout=(60, 5400), stream=True)

    ttfb = time.time() - start
    print(f"Connected: {ttfb:.1f}s, Status: {r.status_code}")

    if r.status_code == 200:
        full = ""
        line_count = 0
        for line in r.iter_lines(decode_unicode=True):
            if not line or not line.startswith("data: "):
                continue
            d = line[6:]
            if d.strip() == "[DONE]":
                break
            try:
                c = json.loads(d)
                delta = c['choices'][0].get('delta', {}).get('content', '')
                full += delta
            except:
                pass

        elapsed = time.time() - start
        content_lines = [l.strip() for l in full.strip().split('\n') if '=' in l]
        print(f"\nDone: {elapsed:.1f}s total")
        print(f"Words output: {len(content_lines)}/{len(words)}")
        print(f"Seconds per word: {elapsed/len(content_lines):.1f}" if content_lines else "No output parsed")
        
        # Show samples
        for l in content_lines[:5]:
            print(f"  {l[:80]}")
        print(f"  ...")
        for l in content_lines[-3:]:
            print(f"  {l[:80]}")
            
        # Estimate for all 3484
        if content_lines:
            total_est = elapsed * (3484 / len(content_lines)) / 3600
            print(f"\nEstimated total time: {total_est:.1f} hours")
    else:
        print(f"HTTP {r.status_code}: {r.text[:300]}")
except Exception as e:
    print(f"Exception ({time.time()-start:.0f}s): {e}")
