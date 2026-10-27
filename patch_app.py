with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

old = '''validate_document_markers(source_text, corrected)
    return clean_thai_ocr_text(corrected)


PAGE_MARKER_RE'''

new = '''validate_document_markers(source_text, corrected)
    return clean_thai_ocr_text(corrected)


@with_api_retry()
def gemini_read_handwriting(image_bytes: bytes, model: str) -> str:
    """ให้ Gemini อ่านลายมือภาษาไทยจากภาพโดยตรง"""
    api_key = gemini_api_key()
    if not api_key:
        raise RuntimeError("ยังไม่ได้ตั้งค่า GEMINI_API_KEY")
    payload = {
        "contents": [{"parts": [
            {"text": HANDWRITING_PROMPT},
            {"inline_data": {"mime_type": "image/png",
                             "data": base64.b64encode(image_bytes).decode("ascii")}},
        ]}],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 8192,
        },
    }
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            result = json.loads(response.read().decode("utf-8"))
        text = "\n".join(
            p["text"] for p in result["candidates"][0]["content"]["parts"] if p.get("text")
        ).strip()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"Gemini HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"เชื่อมต่อ Gemini ไม่สำเร็จ: {exc.reason}") from exc
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("Gemini ไม่ส่งข้อความกลับมา") from exc

    if not text:
        raise RuntimeError("Gemini ส่งข้อความว่างกลับมา")
    if has_repetition_loop(text):
        raise RuntimeError("ผลลัพธ์วนซ้ำผิดปกติ (อาจ hallucinate)")
    return text


PAGE_MARKER_RE'''

if old in content:
    content = content.replace(old, new, 1)
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Successfully inserted gemini_read_handwriting function")
else:
    print("Old text not found")
    # Debug: show what's around line 500
    lines = content.split('\n')
    for i in range(498, 505):
        if i < len(lines):
            print(f"{i+1}: {repr(lines[i])}")
