# Excel Screenshot to PDF Dashboard App

A Streamlit app that turns an uploaded Excel workbook into a screenshot-style PDF.

## What it does

1. Upload an `.xlsx` workbook.
2. Optionally upload a `.pptx` file to reuse slide 1 as the intro page.
3. LibreOffice renders each worksheet to PDF.
4. The app converts each worksheet page into an image.
5. It places each worksheet screenshot onto a PDF page with a right-side BYU Marriott logo bar.
6. The worksheet image itself is not rebuilt or redrawn, so the chart labels and slanted dates stay in the Excel/LibreOffice screenshot style.

## Files to upload to GitHub

- `app.py`
- `requirements.txt`
- `packages.txt`
- `.streamlit/config.toml`
- `assets/default_logo.png`

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

You also need LibreOffice installed locally.

## Streamlit Cloud

`packages.txt` installs LibreOffice on Streamlit Cloud.
