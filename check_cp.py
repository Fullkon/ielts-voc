from pathlib import Path
import json
cp = Path('collocations_checkpoint.json')
if not cp.exists():
    print("No checkpoint yet - first batch still processing")
else:
    d = json.load(open(cp, encoding='utf-8'))
    print(f'Checkpoint found! Words: {len(d)}')
    for w, c in list(d.items())[:5]:
        print(f'  {w}: {c}')
