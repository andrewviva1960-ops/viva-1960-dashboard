with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\VIVA 1960 Dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find the HTML template
tpl_start = content.find("HTML_TEMPLATE = '''")
print(f"Template starts at char {tpl_start}")
if tpl_start >= 0:
    # Find the closing triple-quote
    tpl_body_start = tpl_start + len("HTML_TEMPLATE = '''")
    tpl_end = content.find("'''", tpl_body_start)
    print(f"Template ends at char {tpl_end}")
    template = content[tpl_body_start:tpl_end]
    print(f"Template length: {len(template)} chars")
    
    # Check for triple quotes INSIDE the template
    tq = template.find("'''")
    if tq >= 0:
        print(f"WARNING: Triple quote found inside template at offset {tq}")
        print(f"Context: ...{template[max(0,tq-20):tq+20]}...")
    else:
        print("No triple quotes inside template - OK")
    
    # Extract JS from template
    script_start = template.find('<script>') + 8
    script_end = template.rfind('</script>')
    js = template[script_start:script_end]
    print(f"JS length: {len(js)} chars")
    
    # Check first and last 200 chars of JS
    print(f"\nJS starts with: {js[:200]}")
    print(f"\nJS ends with: {js[-200:]}")
