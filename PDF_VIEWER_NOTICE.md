# PDF viewer implementation note

MyOCR PDF uses Streamlit's built-in `st.pdf` element for the uploaded PDF.
The application passes the original PDF bytes directly to Streamlit instead of
embedding them in nested `data:` URLs or loading a third-party JavaScript CDN.
This avoids browser iframe size limits and external-script failures in deployed
environments.

The UI always includes a direct PDF download button as a fallback. If the
browser cannot display the embedded PDF, users can open the downloaded file in
their normal PDF reader without uploading it again.

The legacy EmbedPDF assets remain in the repository for reference, but they are
not part of the active viewer path.