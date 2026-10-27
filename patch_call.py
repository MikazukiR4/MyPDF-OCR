with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find call sites (not definition)
idx = 0
call_sites = []
while True:
    idx = content.find('extract_pdf(', idx)
    if idx == -1:
        break
    if idx == 0 or content[idx-4:idx] != 'def ':
        call_sites.append(idx)
    idx += 1

print(f"Found {len(call_sites)} call sites:")
for pos in call_sites:
    print(f"  Position {pos}: {repr(content[pos:pos+100])}")
