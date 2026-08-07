import json, os
from pathlib import Path

cp = 'collocations_checkpoint.json'
p = Path(cp)
status = p.exists()
print(f'Checkpoint exists: {status}')
print(f'Size: {p.stat().st_size} bytes' if status else 'N/A')
if status:
    d = json.load(open(cp, encoding='utf-8'))
    print(f'Words processed: {len(d)}')
    for w, c in list(d.items())[:8]:
        print(f'  {w}: {c}')
    print(f'  ...')
    for w, c in list(d.items())[-3:]:
        print(f'  {w}: {c}')
