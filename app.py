import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
from PIL import Image, ImageChops, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

APP_TITLE = "Excel Direct Screenshot Export Builder"
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


def find_executable(names: list[str]) -> str | None:
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


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


def prepare_workbook_for_rendering(source_xlsx: Path, output_xlsx: Path, include_sheets: list[str], margins: float) -> list[str]:
    """Fallback mode only. This touches the workbook with openpyxl, so it is not as faithful as direct render."""
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
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=margins, right=margins, top=margins, bottom=margins, header=0.1, footer=0.1)
    wb.save(output_xlsx)
    return visible_sheets


def convert_to_pdf(input_file: Path, output_dir: Path, soffice_path: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice_path, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(input_file)]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=300)
    if result.returncode != 0:
        raise RuntimeError("LibreOffice failed to convert the file to PDF.\n\n" + result.stderr)
    pdf_path = output_dir / f"{input_file.stem}.pdf"
    if not pdf_path.exists():
        pdfs = list(output_dir.glob("*.pdf"))
        if not pdfs:
            raise FileNotFoundError("LibreOffice did not create a PDF.")
        pdf_path = pdfs[0]
    return pdf_path


def pdf_to_images_poppler(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    pdftoppm = find_executable(["pdftoppm"])
    if not pdftoppm:
        raise RuntimeError("pdftoppm was not found. Add poppler-utils to packages.txt.")
    output_dir.mkdir(parents=True, exist_ok=True)
    prefix = output_dir / "page"
    cmd = [pdftoppm, "-png", "-r", str(dpi), str(pdf_path), str(prefix)]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=300)
    if result.returncode != 0:
        raise RuntimeError("Poppler failed to render the PDF pages.\n\n" + result.stderr)
    return sorted(output_dir.glob("page-*.png"))


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


def transparent_circle_logo(source: Path, output: Path) -> Path:
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
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def fit_text_lines(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> list[str]:
    words = text.split()
    lines, current = [], ""
    for word in words:
        test = word if not current else f"{current} {word}"
        if draw.textbbox((0, 0), test, font=font)[2] <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def make_intro_page(title: str, logo_path: Path | None) -> Image.Image:
    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(page)
    draw.rectangle((0, 0, PAGE_W, 95), fill=TITLE_COLOR)
    draw.text((120, 210), title, fill=TITLE_COLOR, font=load_font(54, bold=True))
    draw.text((120, 300), "Generated from the uploaded Excel workbook.", fill=TEXT_COLOR, font=load_font(28))
    if logo_path and logo_path.exists():
        logo = Image.open(logo_path).convert("RGBA")
        logo.thumbnail((300, 300), Image.LANCZOS)
        page.paste(logo.convert("RGB"), (1450, 150), logo)
    return page


def extract_intro_page(template_pptx: Path, tmpdir: Path, soffice_path: str) -> Image.Image:
    pdf_path = convert_to_pdf(template_pptx, tmpdir / "template_pdf", soffice_path)
    pages = pdf_to_images_poppler(pdf_path, tmpdir / "template_pages", 170)
    if not pages:
        raise RuntimeError("Template PowerPoint rendered no pages.")
    return Image.open(pages[0]).convert("RGB").resize((PAGE_W, PAGE_H), Image.LANCZOS)


def make_sheet_page(sheet_img_path: Path, sheet_name: str, logo_path: Path | None, show_sheet_name: bool) -> Image.Image:
    page = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(page)
    sidebar_x = PAGE_W - SIDEBAR_W
    draw.rectangle((sidebar_x, 0, PAGE_W, PAGE_H), fill=SIDEBAR_BG)
    draw.line((sidebar_x, 20, sidebar_x, PAGE_H - 20), fill=DIVIDER, width=2)

    if logo_path and logo_path.exists():
        logo = Image.open(logo_path).convert("RGBA")
        logo.thumbnail((120, 120), Image.LANCZOS)
        x = sidebar_x + (SIDEBAR_W - logo.width) // 2
        page.paste(logo.convert("RGB"), (x, 35), logo)

    if show_sheet_name:
        font = load_font(18, bold=True)
        label = "Nitty Gritty" if sheet_name == "NittyGrittySheet" else sheet_name
        y = 180
        for line in fit_text_lines(draw, label, font, SIDEBAR_W - 16):
            box = draw.textbbox((0, 0), line, font=font)
            draw.text((sidebar_x + (SIDEBAR_W - (box[2] - box[0])) / 2, y), line, fill=TITLE_COLOR, font=font)
            y += 28

    shot = Image.open(sheet_img_path).convert("RGB")
    scale = min(CONTENT_W / shot.width, CONTENT_H / shot.height)
    shot = shot.resize((int(shot.width * scale), int(shot.height * scale)), Image.LANCZOS)
    left = MARGIN_X + (CONTENT_W - shot.width) // 2
    top = MARGIN_Y + (CONTENT_H - shot.height) // 2
    page.paste(shot, (left, top))
    return page


def build_pdf(pages: list[Image.Image], output_pdf: Path) -> Path:
    rgb_pages = [p.convert("RGB") for p in pages]
    rgb_pages[0].save(output_pdf, "PDF", resolution=180.0, save_all=True, append_images=rgb_pages[1:])
    return output_pdf


def build_powerpoint(pages: list[Image.Image], output_pptx: Path, working_dir: Path) -> Path:
    """Build a PowerPoint using the same full-page screenshot images as the PDF.

    This keeps the worksheet screenshots exactly in the same wrapped format, but
    gives the user a .pptx instead of a .pdf. Each page image becomes one slide.
    """
    prs = Presentation()
    prs.slide_width = 12192000   # 13.333 inches
    prs.slide_height = 6858000   # 7.5 inches
    blank = prs.slide_layouts[6]

    slide_image_dir = working_dir / "pptx_slide_images"
    slide_image_dir.mkdir(parents=True, exist_ok=True)

    for index, page in enumerate(pages, start=1):
        img_path = slide_image_dir / f"slide_{index:02d}.jpg"
        page.convert("RGB").save(img_path, quality=94, optimize=True)
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(str(img_path), 0, 0, width=prs.slide_width, height=prs.slide_height)

    prs.save(output_pptx)
    return output_pptx


def make_zip(files: list[Path], zip_path: Path) -> Path:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for file_path in files:
            archive.write(file_path, arcname=file_path.name)
    return zip_path


def convert_excel_to_export(
    uploaded_excel,
    selected_sheets: list[str],
    all_sheet_names: list[str],
    uploaded_logo,
    uploaded_intro_pptx,
    title: str,
    output_format: str,
    pure_direct_render: bool,
    crop_pages: bool,
    show_sheet_name: bool,
    dpi: int,
    margins: float,
) -> tuple[bytes, bytes, bytes, str, str]:
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
            transparent_circle_logo(raw_logo, logo_path)
        elif DEFAULT_LOGO.exists():
            logo_path = DEFAULT_LOGO

        intro_template = None
        if uploaded_intro_pptx:
            intro_template = save_upload(uploaded_intro_pptx, tmpdir / sanitize_filename(uploaded_intro_pptx.name))

        # TRUE screenshot-style path: do not open or save the workbook with openpyxl.
        # This avoids changing chart XML/cache before LibreOffice renders it.
        if pure_direct_render and selected_sheets == all_sheet_names:
            render_source = excel_path
            rendered_names = all_sheet_names
        else:
            render_source = tmpdir / "prepared.xlsx"
            rendered_names = prepare_workbook_for_rendering(excel_path, render_source, selected_sheets, margins)

        raw_pdf = convert_to_pdf(render_source, tmpdir / "raw_pdf", soffice_path)
        raw_pdf_bytes = raw_pdf.read_bytes()
        page_images = pdf_to_images_poppler(raw_pdf, tmpdir / "raw_pages", dpi)
        if not page_images:
            raise RuntimeError("LibreOffice produced a PDF, but no pages could be rendered.")

        screenshot_paths = []
        for idx, img_path in enumerate(page_images, start=1):
            working = img_path
            if crop_pages:
                cropped = tmpdir / "cropped" / f"page_{idx:02d}.png"
                cropped.parent.mkdir(exist_ok=True)
                working = crop_white_space(working, cropped)
            resized = tmpdir / "resized" / f"page_{idx:02d}.jpg"
            resized.parent.mkdir(exist_ok=True)
            working = resize_for_page(working, resized)
            screenshot_paths.append(working)

        final_pages = []
        if intro_template:
            final_pages.append(extract_intro_page(intro_template, tmpdir, soffice_path))
        else:
            final_pages.append(make_intro_page(title, logo_path))

        # If LibreOffice emits more pages than sheet names, keep all pages and label extras as Page N.
        for i, img_path in enumerate(screenshot_paths):
            label = rendered_names[i] if i < len(rendered_names) else f"Page {i + 1}"
            final_pages.append(make_sheet_page(img_path, label, logo_path, show_sheet_name))

        safe_title = sanitize_filename(title.replace(" ", "_"))

        if output_format == "PowerPoint (.pptx)":
            final_output = tmpdir / f"{safe_title}.pptx"
            build_powerpoint(final_pages, final_output, tmpdir)
            mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        else:
            final_output = tmpdir / f"{safe_title}.pdf"
            build_pdf(final_pages, final_output)
            mime = "application/pdf"

        zip_path = tmpdir / f"{safe_title}.zip"
        make_zip([final_output, raw_pdf], zip_path)
        return final_output.read_bytes(), zip_path.read_bytes(), raw_pdf_bytes, final_output.name, mime


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")
    st.title("📄 Excel Direct Screenshot Export Builder")
    st.caption("Direct-render the uploaded workbook through LibreOffice, then wrap those screenshots into either a PDF or PowerPoint. No chart rebuilding by default.")

    with st.sidebar:
        st.header("Settings")
        title = st.text_input("Output file name", value="Placement_Report_Screenshots")
        output_format = st.radio(
            "Output format",
            options=["PDF", "PowerPoint (.pptx)"],
            index=0,
            help="PDF is best for sharing/printing. PowerPoint gives one full-page screenshot slide per page.",
        )
        pure_direct_render = st.checkbox(
            "Pure screenshot mode: do not edit workbook before rendering",
            value=True,
            help="Best match for Excel chart formatting. If you select only some sheets, the app must prepare a copy and may alter chart internals.",
        )
        crop_pages = st.checkbox("Crop extra white space around screenshots", value=True)
        show_sheet_name = st.checkbox("Show sheet name in the right sidebar", value=True)
        dpi = st.slider("Screenshot resolution", min_value=140, max_value=260, value=190, step=10)
        margins = st.slider("Fallback mode Excel print margins", min_value=0.05, max_value=0.40, value=0.20, step=0.05)

    col1, col2 = st.columns(2)
    with col1:
        excel_upload = st.file_uploader("Upload Excel workbook (.xlsx)", type=["xlsx"])
        logo_upload = st.file_uploader("Optional logo image (.png, .jpg)", type=["png", "jpg", "jpeg"])
    with col2:
        intro_upload = st.file_uploader(
            "Optional PowerPoint template for intro page (.pptx)",
            type=["pptx"],
            help="If uploaded, the app reuses the first slide as the PDF intro page.",
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
    if pure_direct_render and selected_sheets != sheet_names:
        st.warning("Pure screenshot mode can only render the workbook as-is. Because you selected a subset/order, the app will use fallback mode and prepare a copy of the workbook.")
    st.write(f"Selected **{len(selected_sheets)}** sheet(s).")

    if st.button(f"Generate {output_format}", type="primary", disabled=not selected_sheets):
        try:
            with st.spinner(f"Rendering workbook and building {output_format}..."):
                output_bytes, zip_bytes, raw_pdf_bytes, output_name, output_mime = convert_excel_to_export(
                    uploaded_excel=excel_upload,
                    selected_sheets=selected_sheets,
                    all_sheet_names=sheet_names,
                    uploaded_logo=logo_upload,
                    uploaded_intro_pptx=intro_upload,
                    title=title,
                    output_format=output_format,
                    pure_direct_render=pure_direct_render,
                    crop_pages=crop_pages,
                    show_sheet_name=show_sheet_name,
                    dpi=dpi,
                    margins=margins,
                )
            st.success(f"{output_format} created successfully.")
            st.download_button(f"Download final {output_format}", data=output_bytes, file_name=output_name, mime=output_mime)
            st.download_button("Download ZIP", data=zip_bytes, file_name=output_name.rsplit('.', 1)[0] + ".zip", mime="application/zip")
            st.download_button("Download raw LibreOffice PDF for comparison", data=raw_pdf_bytes, file_name="raw_libreoffice_render.pdf", mime="application/pdf")
            st.caption("If the raw LibreOffice PDF does not match Microsoft Excel, the mismatch is coming from LibreOffice chart rendering, not from the PDF/PowerPoint wrapper.")
        except Exception as exc:
            st.error(str(exc))
            st.caption("Make sure packages.txt contains both libreoffice and poppler-utils on Streamlit Cloud.")


if __name__ == "__main__":
    main()
