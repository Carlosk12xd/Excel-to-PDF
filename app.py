import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageChops, ImageDraw, ImageFont

APP_TITLE = "Excel Screenshot to PDF Builder"
DEFAULT_LOGO = Path("assets/default_logo.png")

PAGE_W = 1920
PAGE_H = 1080
SIDEBAR_W = 170
MARGIN_X = 20
MARGIN_Y = 20
CONTENT_W = PAGE_W - SIDEBAR_W - 2 * MARGIN_X - 10
CONTENT_H = PAGE_H - 2 * MARGIN_Y
SIDEBAR_BG = (245, 246, 248)
DIVIDER = (205, 208, 214)
TITLE_COLOR = (16, 49, 101)
TEXT_COLOR = (45, 45, 45)


def flatten_to_rgb(image: Image.Image, background=(255, 255, 255)) -> Image.Image:
    """Return an RGB image, safely flattening transparency onto a white background.

    Pillow's PDF writer expects RGB/CMYK/L images. Uploaded logos and PNGs often
    arrive as RGBA, LA, or P-with-transparency images; passing those through can
    raise errors like "image has wrong mode." This helper normalizes every image
    before PDF export or compositing.
    """
    if image.mode == "RGB":
        return image
    if image.mode in ("RGBA", "LA") or (image.mode == "P" and "transparency" in image.info):
        rgba = image.convert("RGBA")
        bg = Image.new("RGBA", rgba.size, background + (255,))
        bg.alpha_composite(rgba)
        return bg.convert("RGB")
    return image.convert("RGB")


def paste_transparent(base: Image.Image, overlay: Image.Image, xy: tuple[int, int]) -> Image.Image:
    """Safely paste a transparent PNG/logo onto an RGB page.

    This avoids Pillow mode/mask issues by doing the composite in RGBA mode and
    then converting back to RGB for PDF output.
    """
    base_rgba = base.convert("RGBA")
    overlay_rgba = overlay.convert("RGBA")
    base_rgba.alpha_composite(overlay_rgba, dest=(int(xy[0]), int(xy[1])))
    return base_rgba.convert("RGB")


def find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def find_pdftoppm() -> str | None:
    candidates = [
        shutil.which("pdftoppm"),
        r"C:\\Program Files\\poppler\\Library\\bin\\pdftoppm.exe",
        r"C:\\poppler\\Library\\bin\\pdftoppm.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def save_upload(uploaded_file, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(uploaded_file.getbuffer())
    return destination


def sanitize_filename(name: str) -> str:
    allowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_ .()"
    return "".join(ch if ch in allowed else "_" for ch in name).strip() or "file"


def get_sheet_names_from_upload(uploaded_excel) -> list[str]:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded_excel.getbuffer())
        tmp_path = Path(tmp.name)
    try:
        wb = load_workbook(tmp_path, read_only=True)
        return wb.sheetnames
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass


def prepare_workbook_for_rendering(
    source_xlsx: Path,
    output_xlsx: Path,
    include_sheets: list[str],
    fit_each_sheet_to_one_page: bool,
    margins: float,
) -> list[str]:
    wb = load_workbook(source_xlsx)
    visible_sheets = []

    for ws in wb.worksheets:
        if ws.title in include_sheets:
            ws.sheet_state = "visible"
            visible_sheets.append(ws.title)
        else:
            ws.sheet_state = "hidden"

    if not visible_sheets:
        raise ValueError("No sheets selected.")

    for ws in wb.worksheets:
        if ws.title not in visible_sheets:
            continue

        min_r = min_c = 10**9
        max_r = max_c = 0
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    min_r = min(min_r, cell.row)
                    min_c = min(min_c, cell.column)
                    max_r = max(max_r, cell.row)
                    max_c = max(max_c, cell.column)

        if max_r:
            ws.print_area = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"

        ws.page_setup.orientation = "landscape"
        ws.page_margins = PageMargins(
            left=margins,
            right=margins,
            top=margins,
            bottom=margins,
            header=0.1,
            footer=0.1,
        )

        if fit_each_sheet_to_one_page:
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(output_xlsx)
    return visible_sheets


def convert_to_pdf(input_file: Path, output_dir: Path, soffice_path: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice_path,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(input_file),
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=300)
    if result.returncode != 0:
        raise RuntimeError(
            "LibreOffice failed to convert the file to PDF.\n\n"
            f"STDOUT:\n{result.stdout}\n\nSTDERR:\n{result.stderr}"
        )

    pdf_path = output_dir / f"{input_file.stem}.pdf"
    if not pdf_path.exists():
        pdfs = list(output_dir.glob("*.pdf"))
        if not pdfs:
            raise FileNotFoundError("LibreOffice did not create a PDF.")
        pdf_path = pdfs[0]
    return pdf_path


def _page_sort_key(path: Path) -> tuple[int, str]:
    # pdftoppm names files like prefix-1.png, prefix-01.png, or prefix-001.png.
    match = re.search(r"-(\d+)\.png$", path.name)
    return (int(match.group(1)) if match else 10**9, path.name)


def pdf_to_images(pdf_path: Path, output_dir: Path, dpi: int, first_page: int | None = None, last_page: int | None = None) -> list[Path]:
    pdftoppm = find_pdftoppm()
    if not pdftoppm:
        raise RuntimeError(
            "Poppler was not found. Add poppler-utils to packages.txt on Streamlit Cloud, "
            "or install Poppler locally."
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    output_prefix = output_dir / "page"
    cmd = [pdftoppm, "-png", "-r", str(dpi)]
    if first_page is not None:
        cmd.extend(["-f", str(first_page)])
    if last_page is not None:
        cmd.extend(["-l", str(last_page)])
    cmd.extend([str(pdf_path), str(output_prefix)])

    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=300)
    if result.returncode != 0:
        raise RuntimeError(
            "Poppler failed to render PDF pages.\n\n"
            f"STDOUT:\n{result.stdout}\n\nSTDERR:\n{result.stderr}"
        )

    image_paths = sorted(output_dir.glob("page-*.png"), key=_page_sort_key)
    if not image_paths:
        raise FileNotFoundError("Poppler did not create any PNG page images.")
    return image_paths


def crop_white_space(image_path: Path, output_path: Path, threshold: int = 12, margin_px: int = 18) -> Path:
    image = Image.open(image_path).convert("RGB")
    white = Image.new("RGB", image.size, (255, 255, 255))
    diff = ImageChops.difference(image, white)
    mask = diff.point(lambda p: 255 if p > threshold else 0)
    bbox = mask.getbbox()
    if bbox is None:
        image.save(output_path)
        return output_path

    left = max(0, bbox[0] - margin_px)
    top = max(0, bbox[1] - margin_px)
    right = min(image.width, bbox[2] + margin_px)
    bottom = min(image.height, bbox[3] + margin_px)
    image.crop((left, top, right, bottom)).save(output_path)
    return output_path


def resize_for_page(image_path: Path, output_path: Path, max_dimension: int = 2600) -> Path:
    image = Image.open(image_path).convert("RGB")
    scale = min(max_dimension / image.width, max_dimension / image.height, 1.0)
    if scale < 1.0:
        image = image.resize((int(image.width * scale), int(image.height * scale)), Image.LANCZOS)
    image.save(output_path, quality=94, optimize=True)
    return output_path


def ellipse_transparent_logo(source: Path, output: Path) -> Path:
    image = Image.open(source).convert("RGBA")
    mask = Image.new("L", image.size, 0)
    draw = ImageDraw.Draw(mask)
    pad = int(min(image.size) * 0.01)
    draw.ellipse((pad, pad, image.width - pad, image.height - pad), fill=255)
    image.putalpha(mask)
    image.save(output)
    return output


def load_font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def fit_text_lines(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> list[str]:
    words = text.split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        test = f"{current} {word}"
        width = draw.textbbox((0, 0), test, font=font)[2]
        if width <= max_width:
            current = test
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def create_generated_intro_page(title: str, logo_path: Path | None) -> Image.Image:
    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(page)
    draw.rectangle((0, 0, PAGE_W, 95), fill=(16, 49, 101))
    title_font = load_font(54, bold=True)
    subtitle_font = load_font(28)
    draw.text((120, 210), title, fill=TITLE_COLOR, font=title_font)
    draw.text((120, 300), "Generated from the uploaded Excel workbook.", fill=TEXT_COLOR, font=subtitle_font)
    if logo_path and logo_path.exists():
        logo = Image.open(logo_path).convert("RGBA")
        logo.thumbnail((300, 300), Image.LANCZOS)
        page = paste_transparent(page, logo, (1450, 150))
    return flatten_to_rgb(page)


def extract_intro_page(template_pptx: Path, working_dir: Path, soffice_path: str) -> Image.Image:
    pdf_path = convert_to_pdf(template_pptx, working_dir / "template_pdf", soffice_path)
    pages = pdf_to_images(pdf_path, working_dir / "template_pages", dpi=170, first_page=1, last_page=1)
    return flatten_to_rgb(Image.open(pages[0]))


def make_sheet_page(sheet_img_path: Path, sheet_name: str, logo_path: Path | None, show_sheet_name: bool) -> Image.Image:
    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(page)

    sidebar_x = PAGE_W - SIDEBAR_W
    draw.rectangle((sidebar_x, 0, PAGE_W, PAGE_H), fill=SIDEBAR_BG)
    draw.line((sidebar_x, 20, sidebar_x, PAGE_H - 20), fill=DIVIDER, width=2)

    if logo_path and logo_path.exists():
        logo = Image.open(logo_path).convert("RGBA")
        logo.thumbnail((120, 120), Image.LANCZOS)
        logo_x = sidebar_x + (SIDEBAR_W - logo.width) // 2
        page = paste_transparent(page, logo, (logo_x, 35))
        draw = ImageDraw.Draw(page)

    if show_sheet_name:
        font = load_font(18, bold=True)
        text_area_w = SIDEBAR_W - 16
        label = "Nitty Gritty" if sheet_name == "NittyGrittySheet" else sheet_name
        lines = fit_text_lines(draw, label, font, text_area_w)
        y = 180
        for line in lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            line_w = bbox[2] - bbox[0]
            draw.text((sidebar_x + (SIDEBAR_W - line_w) / 2, y), line, fill=TITLE_COLOR, font=font)
            y += 28

    shot = flatten_to_rgb(Image.open(sheet_img_path))
    scale = min(CONTENT_W / shot.width, CONTENT_H / shot.height)
    new_size = (int(shot.width * scale), int(shot.height * scale))
    shot = shot.resize(new_size, Image.LANCZOS)
    left = MARGIN_X + (CONTENT_W - shot.width) // 2
    top = MARGIN_Y + (CONTENT_H - shot.height) // 2
    page.paste(shot, (left, top))
    return page


def build_pdf(page_images: list[Image.Image], output_pdf: Path) -> Path:
    """Save pages as a PDF after forcing every page into RGB mode.

    This fixes Pillow's "image has wrong mode" error caused by transparent PNGs
    or palette images being passed to the PDF writer.
    """
    if not page_images:
        raise ValueError("No pages were generated for the PDF.")
    rgb_pages = [flatten_to_rgb(img) for img in page_images]
    first, rest = rgb_pages[0], rgb_pages[1:]
    first.save(output_pdf, "PDF", resolution=180.0, save_all=True, append_images=rest)
    return output_pdf


def make_zip(file_path: Path, zip_path: Path) -> Path:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(file_path, arcname=file_path.name)
    return zip_path


def convert_excel_to_pdf(
    uploaded_excel,
    selected_sheets: list[str],
    uploaded_logo,
    uploaded_intro_pptx,
    title: str,
    fit_each_sheet_to_one_page: bool,
    crop_pages: bool,
    show_sheet_name: bool,
    dpi: int,
    margins: float,
) -> tuple[bytes, bytes, str]:
    soffice_path = find_soffice()
    if not soffice_path:
        raise RuntimeError("LibreOffice was not found. Install LibreOffice locally or through packages.txt.")

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        excel_path = save_upload(uploaded_excel, tmpdir / sanitize_filename(uploaded_excel.name))

        logo_path = None
        if uploaded_logo:
            raw_logo = save_upload(uploaded_logo, tmpdir / sanitize_filename(uploaded_logo.name))
            logo_path = tmpdir / "logo.png"
            ellipse_transparent_logo(raw_logo, logo_path)
        elif DEFAULT_LOGO.exists():
            logo_path = DEFAULT_LOGO

        intro_template = None
        if uploaded_intro_pptx:
            intro_template = save_upload(uploaded_intro_pptx, tmpdir / sanitize_filename(uploaded_intro_pptx.name))

        prepared_xlsx = tmpdir / "prepared.xlsx"
        rendered_sheet_names = prepare_workbook_for_rendering(
            excel_path,
            prepared_xlsx,
            selected_sheets,
            fit_each_sheet_to_one_page,
            margins,
        )

        pdf_path = convert_to_pdf(prepared_xlsx, tmpdir / "rendered_pdf", soffice_path)
        page_images = pdf_to_images(pdf_path, tmpdir / "sheet_pages", dpi=dpi)
        if len(page_images) < len(rendered_sheet_names):
            raise RuntimeError(f"Expected {len(rendered_sheet_names)} pages, but rendered only {len(page_images)}.")

        screenshot_paths = []
        for idx, img_path in enumerate(page_images[:len(rendered_sheet_names)], start=1):
            working = img_path
            if crop_pages:
                cropped = tmpdir / "cropped" / f"sheet_{idx:02d}.png"
                cropped.parent.mkdir(exist_ok=True)
                working = crop_white_space(working, cropped)
            resized = tmpdir / "resized" / f"sheet_{idx:02d}.jpg"
            resized.parent.mkdir(exist_ok=True)
            working = resize_for_page(working, resized)
            screenshot_paths.append(working)

        final_pages = []
        if intro_template:
            intro = extract_intro_page(intro_template, tmpdir, soffice_path)
            intro = flatten_to_rgb(intro.resize((PAGE_W, PAGE_H), Image.LANCZOS))
            final_pages.append(intro)
        else:
            final_pages.append(create_generated_intro_page(title, logo_path))

        for sheet_name, img_path in zip(rendered_sheet_names, screenshot_paths):
            final_pages.append(make_sheet_page(img_path, sheet_name, logo_path, show_sheet_name))

        safe_title = sanitize_filename(title.replace(" ", "_"))
        final_pdf = tmpdir / f"{safe_title}.pdf"
        build_pdf(final_pages, final_pdf)

        zip_path = tmpdir / f"{safe_title}.zip"
        make_zip(final_pdf, zip_path)
        return final_pdf.read_bytes(), zip_path.read_bytes(), final_pdf.name


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")
    st.title("📄 Excel Screenshot to PDF Builder")
    st.caption("Upload an Excel workbook and export worksheet screenshots into a PDF. No chart rebuilding — charts stay in the Excel-rendered style.")

    with st.sidebar:
        st.header("Settings")
        title = st.text_input("Output file name", value="Placement_Report_Screenshots")
        fit_one_page = st.checkbox("Fit each worksheet to one landscape page", value=True)
        crop_pages = st.checkbox("Crop extra white space around screenshots", value=True)
        show_sheet_name = st.checkbox("Show sheet name in the right sidebar", value=True)
        dpi = st.slider("Screenshot resolution", min_value=140, max_value=260, value=190, step=10)
        margins = st.slider("Excel print margins", min_value=0.05, max_value=0.40, value=0.20, step=0.05)

    col1, col2 = st.columns(2)
    with col1:
        excel_upload = st.file_uploader("Upload Excel workbook (.xlsx)", type=["xlsx"])
        logo_upload = st.file_uploader("Optional logo image (.png, .jpg)", type=["png", "jpg", "jpeg"])
    with col2:
        intro_upload = st.file_uploader(
            "Optional PowerPoint template for intro page (.pptx)",
            type=["pptx"],
            help="If you upload an older deck, the app will reuse page 1 as the intro page.",
        )

    if not excel_upload:
        st.info("Upload an Excel file to begin.")
        return

    try:
        sheet_names = get_sheet_names_from_upload(excel_upload)
    except Exception as exc:
        st.error(f"Could not read the workbook: {exc}")
        return

    st.subheader("Sheets to include")
    selected_sheets = st.multiselect("Choose sheets and order them", options=sheet_names, default=sheet_names)
    st.write(f"Selected **{len(selected_sheets)}** sheet(s).")

    if st.button("Generate PDF", type="primary", disabled=not selected_sheets):
        try:
            with st.spinner("Rendering workbook screenshots and building PDF..."):
                pdf_bytes, zip_bytes, pdf_name = convert_excel_to_pdf(
                    uploaded_excel=excel_upload,
                    selected_sheets=selected_sheets,
                    uploaded_logo=logo_upload,
                    uploaded_intro_pptx=intro_upload,
                    title=title,
                    fit_each_sheet_to_one_page=fit_one_page,
                    crop_pages=crop_pages,
                    show_sheet_name=show_sheet_name,
                    dpi=dpi,
                    margins=margins,
                )
            st.success("PDF created successfully.")
            st.download_button("Download PDF", data=pdf_bytes, file_name=pdf_name, mime="application/pdf")
            st.download_button("Download ZIP", data=zip_bytes, file_name=pdf_name.replace(".pdf", ".zip"), mime="application/zip")
        except Exception as exc:
            st.error(str(exc))
            st.caption("If this is a dependency error, make sure packages.txt includes both libreoffice and poppler-utils. If it mentions image mode, update to this RGB-fixed version of the app.")


if __name__ == "__main__":
    main()
