import urllib.request

with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\VIVA 1960 Dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

tpl_start = content.find("HTML_TEMPLATE = '''") + len("HTML_TEMPLATE = '''")
tpl_end = content.find("'''", tpl_start)
template = content[tpl_start:tpl_end]

req = urllib.request.Request("http://127.0.0.1:8765/")
req.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
resp = urllib.request.urlopen(req, timeout=10)
actual = resp.read().decode('utf-8')

# Find ALL differences
template_js_start = template.find('<script>') + 8
template_js_end = template.rfind('</script>')
tjs = template[template_js_start:template_js_end]

actual_js_start = actual.find('<script>') + 8
actual_js_end = actual.rfind('</script>')
ajs = actual[actual_js_start:actual_js_end]

# Find all diff regions
i = 0
diffs = []
while i < min(len(tjs), len(ajs)):
    if tjs[i] != ajs[i]:
        start = i
        while i < min(len(tjs), len(ajs)) and tjs[i] != ajs[i]:
            i += 1
        diffs.append((start, tjs[start:i], ajs[start:i]))
    else:
        i += 1

print(f"Total diffs: {len(diffs)}")
for pos, t, a in diffs[:10]:
    print(f"  Pos {pos}: T={repr(t[:60])} A={repr(a[:60])}")
