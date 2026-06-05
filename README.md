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

This version automatically extends horizontal Excel chart ranges to the newest date found in the uploaded workbook. It is not hardcoded to 5/15, 5/22, or any specific pull date. If a future workbook includes 5/29, 6/5, or another later weekly pull, the app will extend the chart ranges to that newest workbook date.

The sidebar includes two controls:

- **Auto-extend charts to newest workbook date**: keep this on for normal use.
- **Ignore dates after today**: leave this off unless your workbook contains blank future placeholder columns that should not show yet.


## Fix for stale weekly chart screenshots

This version does more than detect the newest workbook date. It also rebuilds the two top weekly charts from `NittyGrittySheet` before placing the worksheet screenshots into PowerPoint. This fixes the issue where the app detected 5/22 or a newer pull date, but the Excel-rendered screenshot still visually stopped at 5/15.

Keep **Rebuild weekly charts from NittyGrittySheet** turned on for placement dashboard files.


## Force newest weekly pull into charts

This version also supports a rolling weekly-chart window. Keep **Rebuild weekly charts from NittyGrittySheet** turned on and use **Weekly chart window: newest date plus previous pulls**. The app rebuilds the two top weekly charts from workbook data and uses the newest date found in the uploaded file as the rightmost point. For example, if the uploaded workbook includes `6/5/2026`, the rebuilt weekly charts end at `6/5/2026`; if the next upload includes `6/12/2026`, they end at `6/12/2026`.

Recommended setting: `10`, which shows the newest pull plus the previous weekly pulls, roughly the last two months.
