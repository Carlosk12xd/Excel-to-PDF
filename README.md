# Excel Screenshot to PDF App

This Streamlit app renders an uploaded Excel workbook as screenshot-style PDF pages.
It does not rebuild charts. LibreOffice renders the workbook, so chart formatting, slanted dates, fonts, sizes, and tables stay as close as possible to the Excel workbook.

## What changed in this version

- Uses LibreOffice to export Excel/PPTX to PDF.
- Uses Poppler `pdftoppm` to turn PDF pages into images.
- Fixes Pillow `image has wrong mode` errors by flattening all transparent/palette images to RGB before PDF export.
- Supports optional logo upload and optional PowerPoint intro slide.

## GitHub / Streamlit Cloud files

Your repo should include:

```text
app.py
requirements.txt
packages.txt
.streamlit/config.toml
assets/default_logo.png
```

## packages.txt

```text
libreoffice
poppler-utils
```

## requirements.txt

```text
streamlit>=1.36.0
openpyxl>=3.1.2
Pillow>=10.0.0
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

You also need LibreOffice and Poppler installed locally.
