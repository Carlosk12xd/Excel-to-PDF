import io
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageChops, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


APP_TITLE = "Excel to PowerPoint Dashboard Builder"
DEFAULT_LOGO = Path("assets/default_logo.png")

BYU_NAVY = RGBColor(16, 49, 101)
SIDEBAR_BG = RGBColor(245, 246, 248)
SIDEBAR_DIVIDER = RGBColor(205, 208, 214)


def find_soffice() -> str | None:
    """Find the LibreOffice executable."""
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
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


def prepare_workbook_for_rendering(
    source_xlsx: Path,
    output_xlsx: Path,
    include_sheets: list[str],
    fit_each_sheet_to_one_page: bool,
    margins: float,
) -> list[str]:
    """Copy workbook, hide excluded sheets, and set print layout."""
    wb = load_workbook(source_xlsx)

    visible_sheets = []
    for ws in wb.worksheets:
        if ws.title in include_sheets:
            ws.sheet_state = "visible"
            visible_sheets.append(ws.title)
        else:
            ws.sheet_state = "hidden"

    if not visible_sheets:
        raise ValueError("No sheets selected. Select at least one sheet.")

    # Excel requires at least one visible sheet; this is guaranteed above.
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
    """Convert a spreadsheet or PowerPoint to PDF using LibreOffice."""
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
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=180)
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


def pdf_to_images(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    """Render every PDF page to a PNG image."""
    output_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    image_paths = []
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)

    for page_index in range(len(doc)):
        page = doc[page_index]
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        out_path = output_dir / f"page_{page_index + 1:02d}.png"
        pix.save(str(out_path))
        image_paths.append(out_path)
    doc.close()
    return image_paths


def crop_white_space(image_path: Path, output_path: Path, threshold: int = 12, margin_px: int = 18) -> Path:
    """Crop white margins around an image while keeping a small border."""
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
    cropped = image.crop((left, top, right, bottom))
    cropped.save(output_path)
    return output_path


def resize_for_ppt(image_path: Path, output_path: Path, max_dimension: int = 2200) -> Path:
    """Downsize very large screenshots so the PPTX stays reasonably small."""
    image = Image.open(image_path).convert("RGB")
    scale = min(max_dimension / image.width, max_dimension / image.height, 1.0)
    if scale < 1.0:
        new_size = (int(image.width * scale), int(image.height * scale))
        image = image.resize(new_size, Image.LANCZOS)
    image.save(output_path, quality=92, optimize=True)
    return output_path


def make_circular_logo_transparent(source: Path, output: Path) -> Path:
    """Make the area outside the main dark circular logo transparent.

    This works well for the BYU Marriott circular logo. It keeps the white letters
    inside the circle because it uses an ellipse mask instead of removing all white pixels.
    """
    image = Image.open(source).convert("RGBA")
    rgb = image.convert("RGB")
    pixels = rgb.load()

    dark_points = []
    step = max(1, min(image.size) // 500)
    for y in range(0, image.height, step):
        for x in range(0, image.width, step):
            r, g, b = pixels[x, y]
            if b > r + 15 and b > g + 5 and (r + g + b) < 450:
                dark_points.append((x, y))

    if not dark_points:
        image.save(output)
        return output

    xs = [p[0] for p in dark_points]
    ys = [p[1] for p in dark_points]
    bbox = (min(xs), min(ys), max(xs), max(ys))
    pad = int(min(image.size) * 0.02)
    bbox = (
        max(0, bbox[0] - pad),
        max(0, bbox[1] - pad),
        min(image.width - 1, bbox[2] + pad),
        min(image.height - 1, bbox[3] + pad),
    )

    mask = Image.new("L", image.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse(bbox, fill=255)
    image.putalpha(mask)
    image.save(output)
    return output


def add_generated_intro_slide(prs: Presentation, title: str, subtitle: str, logo_path: Path | None) -> None:
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.75))
    bar.fill.solid()
    bar.fill.fore_color.rgb = BYU_NAVY
    bar.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.85), Inches(1.65), Inches(8.2), Inches(1.0))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = BYU_NAVY

    subtitle_box = slide.shapes.add_textbox(Inches(0.85), Inches(2.65), Inches(7.8), Inches(1.0))
    p2 = subtitle_box.text_frame.paragraphs[0]
    p2.text = subtitle
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(45, 45, 45)

    if logo_path and logo_path.exists():
        slide.shapes.add_picture(str(logo_path), Inches(9.75), Inches(1.25), width=Inches(2.25), height=Inches(2.25))


def add_intro_from_template(prs: Presentation, template_pptx: Path, working_dir: Path, soffice_path: str) -> None:
    """Use the first slide from a template deck as a full-slide screenshot."""
    pdf_dir = working_dir / "template_pdf"
    pdf_path = convert_to_pdf(template_pptx, pdf_dir, soffice_path)
    page_images = pdf_to_images(pdf_path, working_dir / "template_pages", dpi=160)
    if not page_images:
        raise RuntimeError("Template PowerPoint did not render any slides.")

    first_slide_img = page_images[0]
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    slide.shapes.add_picture(str(first_slide_img), 0, 0, width=prs.slide_width, height=prs.slide_height)


def build_powerpoint(
    sheet_image_paths: list[Path],
    sheet_names: list[str],
    output_pptx: Path,
    logo_path: Path | None,
    intro_template_pptx: Path | None,
    working_dir: Path,
    soffice_path: str,
    show_sheet_name: bool,
    sidebar_width_inches: float,
    title: str,
) -> Path:
    prs = Presentation()
    prs.slide_width = 12192000  # 13.333 in
    prs.slide_height = 6858000  # 7.5 in
    blank = prs.slide_layouts[6]

    if intro_template_pptx:
        add_intro_from_template(prs, intro_template_pptx, working_dir, soffice_path)
    else:
        add_generated_intro_slide(
            prs,
            title=title,
            subtitle="Generated from the uploaded Excel dashboard.",
            logo_path=logo_path,
        )

    slide_w = 13.333
    slide_h = 7.5
    left_pad = 0.10
    top_pad = 0.27
    right_pad_before_sidebar = 0.15
    content_w = slide_w - sidebar_width_inches - left_pad - right_pad_before_sidebar
    content_h = 6.95

    for sheet_name, img_path in zip(sheet_names, sheet_image_paths):
        slide = prs.slides.add_slide(blank)

        sidebar_left = Inches(slide_w - sidebar_width_inches)
        sidebar_w = Inches(sidebar_width_inches)

        sidebar = slide.shapes.add_shape(1, sidebar_left, 0, sidebar_w, prs.slide_height)
        sidebar.fill.solid()
        sidebar.fill.fore_color.rgb = SIDEBAR_BG
        sidebar.line.fill.background()

        divider = slide.shapes.add_shape(1, sidebar_left, Inches(0.2), 1, Inches(7.0))
        divider.fill.solid()
        divider.fill.fore_color.rgb = SIDEBAR_DIVIDER
        divider.line.fill.background()

        if logo_path and logo_path.exists():
            logo_size = min(sidebar_width_inches - 0.22, 1.05)
            logo_left = slide_w - sidebar_width_inches + (sidebar_width_inches - logo_size) / 2
            slide.shapes.add_picture(
                str(logo_path),
                Inches(logo_left),
                Inches(0.28),
                width=Inches(logo_size),
                height=Inches(logo_size),
            )

        if show_sheet_name:
            box = slide.shapes.add_textbox(
                Inches(slide_w - sidebar_width_inches + 0.04),
                Inches(1.35),
                Inches(sidebar_width_inches - 0.08),
                Inches(4.8),
            )
            p = box.text_frame.paragraphs[0]
            p.text = "Nitty Gritty" if sheet_name == "NittyGrittySheet" else sheet_name
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = BYU_NAVY

        image = Image.open(img_path)
        iw, ih = image.size
        scale = min((content_w * 96) / iw, (content_h * 96) / ih)
        pic_w = Inches(iw / 96 * scale)
        pic_h = Inches(ih / 96 * scale)
        pic_left = Inches(left_pad) + (Inches(content_w) - pic_w) / 2
        pic_top = Inches(top_pad) + (Inches(content_h) - pic_h) / 2

        slide.shapes.add_picture(str(img_path), pic_left, pic_top, width=pic_w, height=pic_h)

    prs.save(output_pptx)
    return output_pptx


def make_zip(file_path: Path, zip_path: Path) -> Path:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(file_path, arcname=file_path.name)
    return zip_path


def convert_excel_to_pptx(
    uploaded_excel,
    selected_sheets: list[str],
    uploaded_logo,
    uploaded_intro_pptx,
    title: str,
    fit_each_sheet_to_one_page: bool,
    crop_pages: bool,
    make_logo_transparent: bool,
    show_sheet_name: bool,
    dpi: int,
    sidebar_width: float,
    margins: float,
) -> tuple[bytes, bytes, str]:
    soffice_path = find_soffice()
    if not soffice_path:
        raise RuntimeError(
            "LibreOffice was not found. Install LibreOffice locally, or deploy with packages.txt on Streamlit Cloud."
        )

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        excel_path = save_upload(uploaded_excel, tmpdir / sanitize_filename(uploaded_excel.name))

        logo_path = None
        if uploaded_logo:
            raw_logo = save_upload(uploaded_logo, tmpdir / sanitize_filename(uploaded_logo.name))
            logo_path = tmpdir / "logo.png"
            if make_logo_transparent:
                make_circular_logo_transparent(raw_logo, logo_path)
            else:
                logo_path = raw_logo
        elif DEFAULT_LOGO.exists():
            logo_path = DEFAULT_LOGO

        intro_template = None
        if uploaded_intro_pptx:
            intro_template = save_upload(uploaded_intro_pptx, tmpdir / sanitize_filename(uploaded_intro_pptx.name))

        prepared_xlsx = tmpdir / "prepared_dashboard.xlsx"
        rendered_sheet_names = prepare_workbook_for_rendering(
            excel_path,
            prepared_xlsx,
            selected_sheets,
            fit_each_sheet_to_one_page,
            margins,
        )

        pdf_path = convert_to_pdf(prepared_xlsx, tmpdir / "pdf", soffice_path)
        page_images = pdf_to_images(pdf_path, tmpdir / "pages", dpi=dpi)

        if len(page_images) < len(rendered_sheet_names):
            raise RuntimeError(
                f"Expected at least {len(rendered_sheet_names)} rendered pages, but got {len(page_images)}. "
                "Try selecting fewer sheets or disabling one-page fitting."
            )

        final_images = []
        for index, img_path in enumerate(page_images[: len(rendered_sheet_names)], start=1):
            working = img_path
            if crop_pages:
                cropped = tmpdir / "cropped" / f"sheet_{index:02d}.png"
                cropped.parent.mkdir(exist_ok=True)
                working = crop_white_space(working, cropped)

            optimized = tmpdir / "optimized" / f"sheet_{index:02d}.jpg"
            optimized.parent.mkdir(exist_ok=True)
            working = resize_for_ppt(working, optimized)
            final_images.append(working)

        safe_title = sanitize_filename(title.replace(" ", "_"))
        pptx_path = tmpdir / f"{safe_title}.pptx"
        build_powerpoint(
            final_images,
            rendered_sheet_names,
            pptx_path,
            logo_path,
            intro_template,
            tmpdir,
            soffice_path,
            show_sheet_name,
            sidebar_width,
            title,
        )

        zip_path = tmpdir / f"{safe_title}.zip"
        make_zip(pptx_path, zip_path)

        return pptx_path.read_bytes(), zip_path.read_bytes(), pptx_path.name


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


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")

    st.title("📊 Excel to PowerPoint Dashboard Builder")
    st.caption("Upload an Excel dashboard and turn every worksheet into a clean PowerPoint slide.")

    with st.sidebar:
        st.header("Settings")
        title = st.text_input("Output deck name", value="Placement_Report_Updated")
        fit_one_page = st.checkbox("Fit each worksheet to one landscape page", value=True)
        crop_pages = st.checkbox("Crop extra white space around screenshots", value=True)
        show_sheet_name = st.checkbox("Show sheet name in the right sidebar", value=True)
        make_logo_transparent = st.checkbox("Remove outside background from circular logo", value=True)
        dpi = st.slider("Screenshot resolution", min_value=120, max_value=260, value=190, step=10)
        sidebar_width = st.slider("Right logo sidebar width", min_value=0.8, max_value=1.6, value=1.2, step=0.05)
        margins = st.slider("Excel print margins", min_value=0.05, max_value=0.40, value=0.20, step=0.05)

    col1, col2 = st.columns(2)
    with col1:
        excel_upload = st.file_uploader("Upload Excel workbook (.xlsx)", type=["xlsx"])
        logo_upload = st.file_uploader("Optional logo image (.png, .jpg)", type=["png", "jpg", "jpeg"])
    with col2:
        intro_upload = st.file_uploader(
            "Optional PowerPoint template for intro slide (.pptx)",
            type=["pptx"],
            help="If you upload your old deck, the app will use slide 1 as the intro slide.",
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
    default_sheets = sheet_names
    selected_sheets = st.multiselect(
        "Choose sheets and order them",
        options=sheet_names,
        default=default_sheets,
    )

    st.write(f"Selected **{len(selected_sheets)}** sheet(s).")

    if st.button("Generate PowerPoint", type="primary", disabled=not selected_sheets):
        try:
            with st.spinner("Rendering Excel sheets and building PowerPoint..."):
                pptx_bytes, zip_bytes, pptx_name = convert_excel_to_pptx(
                    uploaded_excel=excel_upload,
                    selected_sheets=selected_sheets,
                    uploaded_logo=logo_upload,
                    uploaded_intro_pptx=intro_upload,
                    title=title,
                    fit_each_sheet_to_one_page=fit_one_page,
                    crop_pages=crop_pages,
                    make_logo_transparent=make_logo_transparent,
                    show_sheet_name=show_sheet_name,
                    dpi=dpi,
                    sidebar_width=sidebar_width,
                    margins=margins,
                )

            st.success("PowerPoint created successfully.")
            st.download_button(
                "Download PowerPoint",
                data=pptx_bytes,
                file_name=pptx_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.download_button(
                "Download ZIP",
                data=zip_bytes,
                file_name=pptx_name.replace(".pptx", ".zip"),
                mime="application/zip",
            )
        except Exception as exc:
            st.error(str(exc))
            st.caption(
                "Common fix: make sure LibreOffice is installed locally, or include packages.txt when deploying to Streamlit Cloud."
            )


if __name__ == "__main__":
    main()
