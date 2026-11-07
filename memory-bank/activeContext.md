# Active Context — MyOCR PDF

## Current Focus (11/7/2026)
Phase 3: Advanced Features — IN PROGRESS 🚧

## Completed Changes

### Phase 1: Quick Wins (Previously completed)
- [x] Fix 1: CSS dark theme consistency in inject_css()
- [x] Fix 2: Visual feedback badge for page selection (st.success with check emoji)
- [x] Fix 3: Skeleton loading spinner in PDF viewer (render_pdf_viewer())
- [x] Fix 4: ETA calculation in progress bar (_ProgressWithETA wrapper)

### Phase 2: Enhanced UX
- [x] Undo/Redo Support: Button + Ctrl+Z keyboard shortcut in sidebar, session state stack (max 50 entries), tracks text changes via _prev_ocr_editor comparison. Undo handler triggers st.rerun() on button press. Redo stack tracked but not yet exposed as UI element.
- [x] OCR Stats Banner: Shows page count, character count, word count, line count after OCR completes — styled HTML banner with badges displayed via ocr_preview.markdown().
- [x] Multi-format Export: Added TXT, HTML, Markdown export options alongside existing DOCX. Format selector dropdown in download section. Helper functions: export_txt_file(), export_html_file(), export_markdown_file() all cached with @st.cache_data.
- [x] Keyboard Shortcuts Help: Sidebar section listing Ctrl+Z (undo), Ctrl+S (download), Esc (cancel) with styled kbd elements.
- [x] PDF Viewer Zoom Note: Caption below PDF viewer explaining + / - zoom controls.
- [x] CSS Additions: .stats-banner, .stats-badge, .stats-value, .shortcut-item, .shortcut-key styles in inject_css().

### Phase 3: Advanced Features (IN PROGRESS)
- [x] OCR Confidence Score: `_calculate_ocr_confidence(text)` function that scores OCR quality 0-100% using heuristics (Thai char ratio, text length, error markers, line structure). Returns overall score, per-page breakdown, and colored badge HTML.
- [x] Auto Language Detection: `_detect_language(text)` function using langdetect library with fallback heuristic. Returns "th", "en", "mixed", or "unknown".
- [x] Enhanced Stats Banner: `format_stats_banner()` now accepts optional confidence and language params, displays them inline with stats.
- [x] Theme Toggle: Radio buttons in sidebar for dark/light mode toggle. Light theme CSS injected via `inject_theme_override_css()` function. Session state persisted in `theme_mode`.
- [x] Batch Processing UI: File uploader switches to `accept_multiple_files=True` when batch mode enabled. Shows batch queue with per-file status badges (pending/processing/done/error). Per-file result tabs with OCR text editor, confidence display, and individual download buttons.
- [x] Batch Processing Execution: `_batch_run_requested` triggers sequential OCR on all uploaded files. Results stored in `batch_results` session state with confidence and language detection. ZIP download support for all results.
- [x] CSS Additions: `.batch-queue`, `.batch-file-item`, `.batch-file-status`, `.status-*`, `.per-page-confidence`, `.page-conf-*`, `.theme-switch-container` styles.

## Files Modified
- app.py (~1930 lines, was ~1575 after Phase 2):
  - Lines 50-67: Added batch processing, theme toggle, and language detection session state
  - Lines 259-433: Added `_calculate_ocr_confidence()`, `_detect_language()`, updated `format_stats_banner()` with confidence/lang support
  - Lines 1323-1380: Added `inject_theme_override_css()` for light theme + Phase 3 component CSS
  - Lines 1383-1384: Call inject_theme_override_css based on theme_mode
  - Lines 1437-1469: Added Theme Toggle radio and Batch Mode checkbox in sidebar
  - Lines 1617-1663: Modified file uploader to support both single-file and batch mode
  - Lines 1679-1763: Added Batch Processing Queue UI with per-file tabs, stats, confidence display
  - Lines 1836-1933: Added Batch Processing Execution logic and ZIP download handler

## Known Issues
- Undo button uses session state key polling (_hidden_undo_btn) — works within Streamlit constraints but not ideal
- Redo functionality tracked in ocr_redo_stack but not yet exposed as UI element (planned for Phase 4)
- _prev_ocr_editor tracking may accumulate redundant entries on rapid edits (FIFO limit of 50 mitigates this)
- Theme toggle works via CSS injection override — Streamlit native theming still uses dark config
- Batch processing runs sequentially (not parallel) due to Streamlit rerun model limitations
