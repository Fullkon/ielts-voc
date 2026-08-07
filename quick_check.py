from pathlib import Path
import json

cp = Path('collocations_checkpoint.json')
print(f'Exists: {cp.exists()}, Size: {cp.stat().st_size if cp.exists() else "N/A"}')
if cp.exists():
    d = json.load(open(cp, encoding='utf-8'))
    print(f'Words: {len(d)}')
    items = list(d.items())
    if items:
        for w, c in items[:5]:
            print(f'  {w}: {c}')
        print(f'  ...')
        for w, c in items[-3:]:
            print(f'  {w}: {c}')
