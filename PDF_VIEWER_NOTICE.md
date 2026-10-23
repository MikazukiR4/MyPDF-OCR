# PDF viewer third-party notices

This project includes the EmbedPDF viewer integration based on:

https://github.com/embedpdf/embed-pdf-viewer

The browser module is loaded from the jsDelivr package:

https://cdn.jsdelivr.net/npm/@embedpdf/snippet@2/dist/embedpdf.js

The integration creates a temporary browser object URL from the selected PDF
bytes and passes it to EmbedPDF. The PDF is not uploaded a second time by the
viewer integration.