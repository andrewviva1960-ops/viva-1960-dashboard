import requests, re, json

r = requests.get('https://www.cbe.org.eg/en/economic-research/statistics/cbe-exchange-rates', timeout=30)
html = r.text

# Find all numeric patterns near currency codes
currencies_needed = ['USD', 'EUR', 'GBP', 'CHF', 'KWD', 'SAR', 'BHD', 'OMR', 'JOD']
results = {}

for cur in currencies_needed:
    # Search for patterns like "USD 48.5" or "USD48.5" or "USD = 48.5"
    pattern = re.findall(rf'{cur}\s*:?\s*([0-9]+\.[0-9]+)', html)
    if pattern:
        results[cur] = [float(x) for x in pattern]

print("=== Direct patterns ===")
for k, v in results.items():
    print(f"{k}: {v}")

# Also search for any table rows that might contain these currencies
# Look for the exchange rates table
table_sections = re.findall(r'<tr[^>]*>.*?</tr>', html, re.DOTALL)
print(f"\n=== Found {len(table_sections)} table rows ===")

for tr in table_sections:
    if any(c in tr for c in currencies_needed):
        # Clean HTML tags
        text = re.sub(r'<[^>]+>', ' ', tr)
        text = re.sub(r'\s+', ' ', text).strip()
        print(text)

# Try to find JSON data
json_patterns = re.findall(r'\[.*?\]', html, re.DOTALL)
for jp in json_patterns[:10]:
    if any(c in jp for c in currencies_needed):
        print(f"\nJSON-like: {jp[:500]}")

# Search for API calls
api_patterns = re.findall(r'https?://[^"\'\\s]+(?:rate|exchange|currency)[^"\'\\s]*', html)
print(f"\n=== API endpoints ({len(api_patterns)}) ===")
for ap in api_patterns[:10]:
    print(ap)
