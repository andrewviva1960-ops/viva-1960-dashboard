import urllib.request

req = urllib.request.Request("http://127.0.0.1:8765/")
req.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
resp = urllib.request.urlopen(req, timeout=10)
html = resp.read().decode('utf-8')

js_start = html.find('<script>') + 8
js_end = html.rfind('</script>')
js = html[js_start:js_end]

# Check for unclosed single-quoted strings (newline inside string = broken)
lines = js.split('\n')
for i, line in enumerate(lines, 1):
    count = 0
    for ch in line:
        if ch == "'":
            count += 1
    if count % 2 != 0:
        print(f'ODD QUOTES line {i}: {line.strip()[:120]}')
