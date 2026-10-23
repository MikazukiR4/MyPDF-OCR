# PDF viewer third-party notices

This project includes the EmbedPDF viewer integration based on:

https://github.com/embedpdf/embed-pdf-viewer

The browser module is loaded from the jsDelivr package:

https://cdn.jsdelivr.net/npm/@embedpdf/snippet@2/dist/embedpdf.js

The integration creates a temporary browser object URL from the selected PDF
bytes and passes it to EmbedPDF. The PDF is not uploaded a second time by the
viewer integration.

## Streamlit iframe note

The integration intentionally uses `st.components.v1.html` because the viewer
contains an ES module and must execute inside an iframe. Streamlit currently
marks this API as deprecated, but `st.html` is not a drop-in replacement: it is
not iframed and JavaScript is disabled unless explicitly enabled. Revisit this
when Streamlit provides an iframe-capable replacement that supports the viewer.

The viewer HTML is repository-owned. Only the uploaded PDF bytes are injected
into its configuration; user-provided text is not treated as HTML or script.