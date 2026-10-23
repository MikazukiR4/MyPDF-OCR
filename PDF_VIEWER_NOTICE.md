# PDF viewer third-party notices

This project includes the single-file PDF viewer from:

https://github.com/MichalAFerber/pdf-viewer.us

The viewer bundles PDF.js v3.11.174 inline under the Apache License 2.0.
The original PDF.js license banners are retained in `pdf-viewer.html`.

The viewer source is adapted only to expose a local host function named
`window.__pdfViewerReadFile` so THSummarizer can load the selected PDF in the
embedded viewer without uploading it a second time.