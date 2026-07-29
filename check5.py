with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\VIVA 1960 Dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Extract JS from the Python file directly
tpl_start = content.find("HTML_TEMPLATE = '''") + len("HTML_TEMPLATE = '''")
tpl_end = content.find("'''", tpl_start)
template = content[tpl_start:tpl_end]

js_start = template.find('<script>') + 8
js_end = template.rfind('</script>')
template_js = template[js_start:js_end]

# Check for render_template_string breaking JS
# Look for what render_template_string does with {{ }}
# It uses Jinja2 which treats {{ }} as variable expressions
# Check if any JS line has {{ or {% 
for i, line in enumerate(template_js.split('\n'), 1):
    if '{{' in line or '{%' in line:
        print(f'JINJA RISK line {i}: {line.strip()[:100]}')
    if '}}' in line:
        print(f'}}} found line {i}: {line.strip()[:100]}')

# Now check if AED or SAR or anything has issues
print()
print("Checking for potential Jinja2 conflicts...")
for i, line in enumerate(template_js.split('\n'), 1):
    s = line.strip()
    if s.startswith('//'):
        continue
    # Look for any {{ that could be Jinja
    # But also check for things like AED that look like HTML tags
    if '<AED' in line or '<SAR' in line or '<EGP' in line:
        print(f'Potential HTML tag conflict line {i}: {s[:100]}')
