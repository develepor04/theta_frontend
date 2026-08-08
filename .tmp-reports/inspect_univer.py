import re
from pathlib import Path

p = Path('node_modules/@univerjs/design/lib/index.js')
text = p.read_text(encoding='utf-8', errors='ignore')
for m in re.finditer(r'function Input\b|const Input\s*=|Input\s*=\s*\(', text):
    start = max(0, m.start() - 40)
    end = min(len(text), m.start() + 500)
    print('MATCH', m.group(), 'at', m.start())
    print(text[start:end][:500])
    print('----')

# Find setValue in facade packages
roots = [
    Path('node_modules/@univerjs/sheets/lib'),
    Path('node_modules/@univerjs/presets/lib'),
]
for root in roots:
    for f in root.rglob('*.js'):
        t = f.read_text(encoding='utf-8', errors='ignore')
        if 'setValues' in t and 'getRange' in t:
            print('FILE', f)
            for m in re.finditer(r'.{0,80}setValues.{0,120}', t):
                print(m.group().replace('\n', ' ')[:200])
                break
            break
