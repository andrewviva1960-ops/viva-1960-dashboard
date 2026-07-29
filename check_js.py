import re
with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\VIVA 1960 Dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Check for Jinja2 template vars
pos = 0
found = False
while True:
    p = content.find('{{', pos)
    if p == -1: break
    e = content.find('}}', p)
    print(f'Jinja var at pos {p}: {content[p:e+2][:80]}')
    pos = p + 2
    found = True
if not found:
    print('No Jinja template vars found')

# Extract JS and check for specific syntax errors
start = content.find('<script>') + 8
end = content.rfind('</script>')
js = content[start:end]

# Check for unescaped backticks
backtick_count = js.count('`')
print(f'Backticks in JS: {backtick_count}')

# Check for common issues - lines with syntax problems
lines = js.split('\n')
for i, line in enumerate(lines, 1):
    s = line.strip()
    # Skip comments
    if s.startswith('//'):
        continue
    # Check for single-quote inside single-quoted string
    if "btoa('" in s:
        print(f'Line {i} btoa call: {s[:100]}')
