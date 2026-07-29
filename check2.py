with open(r'C:\Users\Andro\Downloads\VIVA 1960 Dashboard\check.js', 'r', encoding='utf-8') as f:
    c = f.read()
print('Triple single quotes:', c.count("'''"))
# Check for any characters that might break Python triple-quoted strings
for i, ch in enumerate(c):
    if ord(ch) > 127:
        ctx = c[max(0,i-10):i+10]
        print(f'Non-ASCII at pos {i}: ord={ord(ch)} ctx={repr(ctx)}')
