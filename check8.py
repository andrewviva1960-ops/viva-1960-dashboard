with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\VIVA 1960 Dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find the HTML template
marker = "HTML_TEMPLATE = '''"
tpl_start = content.find(marker) + len(marker)
tpl_end = content.find("'''", tpl_start)
template = content[tpl_start:tpl_end]

# Find JS section
js_start = template.find('<script>') + 8
js_end = template.rfind('</script>')
js = template[js_start:js_end]

# After Python processes the source, check for actual newlines inside single-quoted strings
lines = js.split('\n')
for i, line in enumerate(lines, 1):
    # Check if this line has odd single quotes (string continues on next line)
    count = 0
    in_dq = False
    in_sq = False
    for ch in line:
        if ch == "'" and not in_dq:
            in_sq = not in_sq
        if ch == '"' and not in_sq:
            in_dq = not in_dq
    if in_sq:
        # String is unclosed - it continues to the next line = BROKEN
        print(f'BROKEN single-quoted string at line {i}: {line.strip()[:120]}')
