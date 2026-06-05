# Excel Direct Render Export App

This version is designed to behave like a screenshot workflow.

It sends the uploaded workbook directly to LibreOffice first, then turns that rendered output into page images. The app can wrap those images into either:

- a PDF, or
- a PowerPoint deck with one full-page screenshot slide per page.

Important: it does **not** rebuild or redraw charts in pure screenshot mode. This keeps the output as close as possible to what LibreOffice renders from the workbook.

## Deploy files

Upload these to GitHub:

- `app.py`
- `requirements.txt`
- `packages.txt`
- `.streamlit/config.toml`
- `assets/default_logo.png`

## Streamlit Cloud packages

`packages.txt` must include:

```txt
libreoffice
poppler-utils
```

## Notes

Use the sidebar **Output format** control to choose **PDF** or **PowerPoint (.pptx)**.

If the raw LibreOffice PDF still does not match Microsoft Excel, the issue is LibreOffice's chart compatibility, not the PDF/PowerPoint wrapping. In that case the only exact-match options are to export from Microsoft Excel itself, or upload a PDF that Excel already exported.
