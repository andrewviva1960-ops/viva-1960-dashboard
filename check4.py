with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\check.js', 'r', encoding='utf-8') as f:
    js = f.read()

stack = []
pairs = {')':'(', ']':'[', '}':'{'}
openers = set('([{')
closers = set(')]}')
in_string = False
string_char = None
escaped = False
line = 1
for i, ch in enumerate(js):
    if ch == '\n':
        line += 1
    if escaped:
        escaped = False
        continue
    if ch == '\\':
        escaped = True
        continue
    if in_string:
        if ch == string_char:
            in_string = False
        continue
    if ch in ("'", '"', '`'):
        in_string = True
        string_char = ch
        continue
    if ch in openers:
        stack.append((ch, line))
    elif ch in closers:
        if not stack or stack[-1][0] != pairs[ch]:
            print(f'Mismatch at line {line}: got {ch}, expected {pairs.get(ch, "?")} (stack top: {stack[-1] if stack else "empty"})')
            break
        stack.pop()
if stack:
    print(f'Unclosed at end: {stack[-5:]}')
else:
    print('All brackets/parens/braces matched correctly')
