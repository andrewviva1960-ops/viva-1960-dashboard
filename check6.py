import re

with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\VIVA 1960 Dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Extract HTML template
tpl_start = content.find("HTML_TEMPLATE = '''") + len("HTML_TEMPLATE = '''")
tpl_end = content.find("'''", tpl_start)
template = content[tpl_start:tpl_end]

# Check: does render_template_string have issues with any chars?
# The key issue: Jinja2 auto-escapes or processes {{ }}
# But we already confirmed no {{ in JS

# Let's check: what does the rendered HTML look like for the script tag?
# Fetch the actual page and compare
import urllib.request
import base64

req = urllib.request.Request("http://127.0.0.1:8765/")
req.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
try:
    resp = urllib.request.urlopen(req, timeout=10)
    actual_html = resp.read().decode('utf-8')
    
    # Extract JS from actual HTML
    actual_js_start = actual_html.find('<script>') + 8
    actual_js_end = actual_html.rfind('</script>')
    actual_js = actual_html[actual_js_start:actual_js_end]
    
    # Extract JS from template
    template_js_start = template.find('<script>') + 8
    template_js_end = template.rfind('</script>')
    template_js = template[template_js_start:template_js_end]
    
    print(f"Template JS length: {len(template_js)}")
    print(f"Actual JS length: {len(actual_js)}")
    
    if len(template_js) != len(actual_js):
        print("JS LENGTH MISMATCH!")
        # Find first difference
        for i in range(min(len(template_js), len(actual_js))):
            if template_js[i] != actual_js[i]:
                print(f"First diff at char {i}")
                print(f"  Template: {repr(template_js[max(0,i-20):i+20])}")
                print(f"  Actual:   {repr(actual_js[max(0,i-20):i+20])}")
                break
    else:
        if template_js == actual_js:
            print("JS is IDENTICAL - no corruption")
        else:
            print("JS same length but different content!")
            for i in range(len(template_js)):
                if template_js[i] != actual_js[i]:
                    print(f"First diff at char {i}")
                    print(f"  Template: {repr(template_js[max(0,i-20):i+20])}")
                    print(f"  Actual:   {repr(actual_js[max(0,i-20):i+20])}")
                    break
except Exception as e:
    print(f"Failed to fetch page: {e}")
