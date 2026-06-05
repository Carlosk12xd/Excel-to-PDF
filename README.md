# Excel to PowerPoint Dashboard App

A Streamlit app that converts an uploaded Excel workbook into a PowerPoint deck.

What it does:

1. User uploads an `.xlsx` file.
2. The app sets each worksheet to fit on one landscape page.
3. LibreOffice renders the workbook to PDF.
4. The app converts each PDF page into a clean screenshot.
5. It builds a PowerPoint with one slide per worksheet.
6. It keeps an optional first slide from an uploaded PowerPoint template.
7. It places the BYU Marriott logo in a right sidebar without blocking the data.
8. For placement dashboard files, it can rebuild the two top weekly charts from `NittyGrittySheet` so the newest workbook pull date is included.

## Local setup

```bash
pip install -r requirements.txt
```

You also need LibreOffice installed:

- macOS: install LibreOffice from libreoffice.org, or with Homebrew: `brew install --cask libreoffice`
- Windows: install LibreOffice and make sure `soffice.exe` is on PATH
- Linux/Streamlit Cloud: `packages.txt` installs it

Run:

```bash
streamlit run app.py
```

## Streamlit Cloud deployment

Upload these files to GitHub:

- `app.py`
- `requirements.txt`
- `packages.txt`
- `.streamlit/config.toml`
- `assets/default_logo.png`

Then deploy the repo on Streamlit Cloud.

## Notes

- For best results, use `.xlsx` files with dashboard sheets already formatted similarly to your placement report workbook.
- Uploading a PowerPoint template is optional. When uploaded, the app uses the first slide as the intro slide.
- If you do not upload a logo, the app uses `assets/default_logo.png`.

## Future-date chart handling

This version is not hardcoded to 5/15, 5/22, 6/5, or any specific pull date. It scans the uploaded workbook and uses the newest weekly pull date found in the data.

For BYU Marriott placement dashboards, keep **Rebuild weekly charts from NittyGrittySheet** turned on. The app rebuilds the weekly placement and weekly search-status charts directly from workbook data instead of relying on stale Excel chart rendering.

The rebuilt weekly charts now match the older dashboard style more closely:

- full timeline from the original starting date to the newest workbook date
- x-axis labels placed only on real weekly data points
- first/original date always included
- newest workbook date always included
- regular data-point labels in between, similar to the older chart format
- smaller Excel-style text so labels do not block the chart

Recommended setting: **Weekly x-axis date labels = 18**.
