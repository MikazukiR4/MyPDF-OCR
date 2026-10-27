with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find where extract_pdf is called
old_call = '''text, method, total_pages = extract_pdf(
        uploaded_bytes,'''

new_call = '''text, method, total_pages = extract_pdf(
        uploaded_bytes,
        language,
        mode,
        psm,
        oem,
        engine,
        use_gemini_correction,
        gemini_model,
        progress,
        page_callback,
        pages_to_process,
        page_cache,
        handwriting_mode=handwriting_mode,
    '''

# Check if old_call exists (might have slight variations)
if old_call in content:
    content = content.replace(old_call, new_call, 1)
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Successfully updated extract_pdf call with handwriting_mode")
else:
    print("Old call pattern not found, searching for alternatives...")
    import re
    matches = list(re.finditer(r'extract_pdf\\(\\s*uploaded_bytes', content))
    for m in matches:
        print(f"Found at position {m.start()}: {repr(content[m.start():m.start()+100])}")
