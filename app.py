import io
import zipfile
from dataclasses import dataclass

import streamlit as st
from docx import Document
from PIL import Image, ImageDraw, ImageFont

import fitz
import arabic_reshaper
from bidi.algorithm import get_display


# ----------------------------
# Helpers
# ----------------------------
def render_first_page(pdf_bytes: bytes, dpi=200):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]

    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    pix = page.get_pixmap(matrix=matrix, alpha=False)
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def load_template(file, dpi=200):
    data = file.read()

    if file.name.lower().endswith(".pdf"):
        return render_first_page(data, dpi=dpi)

    return Image.open(io.BytesIO(data)).convert("RGB")


# ----------------------------
# استخراج الاسم + التاريخ
# ----------------------------
def extract_data_from_docx(docx_file):
    data = docx_file.read()
    doc = Document(io.BytesIO(data))

    results = []

    for table in doc.tables:
        if len(table.rows) < 2:
            continue

        headers = [cell.text.strip() for cell in table.rows[0].cells]

        name_col = None
        date_col = None

        for i, h in enumerate(headers):
            h_clean = h.replace(" ", "")
            if h_clean == "الاسم":
                name_col = i
            if "تاريخ" in h_clean:
                date_col = i

        if name_col is not None:
            current_date = ""

            for r in table.rows[1:]:
                name = r.cells[name_col].text.strip() if name_col < len(r.cells) else ""
                date = ""

                if date_col is not None and date_col < len(r.cells):
                    date = r.cells[date_col].text.strip()

                if date:
                    current_date = date

                clean_name = name.replace(" ", "").strip()

                if clean_name and clean_name != "الاسم":
                    results.append((name, current_date))

    return results


def shape_arabic(text):
    text = str(text).strip()

    if not text:
        return ""

    return get_display(arabic_reshaper.reshape(text))


def get_font(font_path, size):
    return ImageFont.truetype(font_path, size)


def fit_font(draw, text, font_path, max_width, size, min_size=30):
    while size > min_size:
        font = get_font(font_path, size)
        bbox = draw.textbbox((0, 0), text, font=font)
        w = bbox[2] - bbox[0]

        if w <= max_width:
            return font

        size -= 2

    return get_font(font_path, min_size)


@dataclass
class Placement:
    x: float
    y: float
    max_w: float


# ----------------------------
# رسم الاسم + التاريخ
# ----------------------------
def draw_on_template(
    img,
    name,
    date,
    font_path,
    font_size,
    p_name,
    p_date,
    color,
    dpi=200
):
    img = img.copy()
    draw = ImageDraw.Draw(img)

    W, H = img.size

    scale = dpi / 72.0
    scaled_font_size = max(12, int(font_size * scale))

    # الاسم
    name = shape_arabic(name)
    x_name = int(W * p_name.x)
    y_name = int(H * p_name.y)
    max_w = int(W * p_name.max_w)

    font_name = fit_font(
        draw,
        name,
        font_path,
        max_w,
        scaled_font_size,
        min_size=max(20, int(scaled_font_size * 0.45))
    )

    draw.text(
        (x_name, y_name),
        name,
        font=font_name,
        fill=color,
        anchor="rm"
    )

    # التاريخ
    date = shape_arabic(date)
    x_date = int(W * p_date.x)
    y_date = int(H * p_date.y)

    font_date = get_font(font_path, max(10, int(scaled_font_size * 0.7)))

    draw.text(
        (x_date, y_date),
        date,
        font=font_date,
        fill=color,
        anchor="rm"
    )

    return img


# ----------------------------
# تحويل الصورة إلى PDF / JPEG
# ----------------------------
def image_to_pdf_bytes(img, dpi=200):
    buf = io.BytesIO()
    rgb_img = img.convert("RGB")
    rgb_img.save(buf, format="PDF", resolution=dpi)
    return buf.getvalue()


def image_to_jpeg_bytes(img, quality=95):
    buf = io.BytesIO()
    rgb_img = img.convert("RGB")
    rgb_img.save(
        buf,
        format="JPEG",
        quality=quality,
        optimize=True,
        progressive=True
    )
    return buf.getvalue()


def safe_filename(text, fallback="certificate"):
    name = "".join(c for c in str(text) if c not in '\\/:*?"<>|').strip()
    return name or fallback


# ----------------------------
# UI
# ----------------------------
st.set_page_config(layout="wide")
st.title("مولد الشهادات ZIP 🔥")

col1, col2 = st.columns(2)

with col1:
    template_file = st.file_uploader("ارفع الشهادة", type=["pdf", "png", "jpg", "jpeg"])
    names_file = st.file_uploader("ارفع ملف الأسماء", type=["docx"])

with col2:
    font_path = st.text_input("مسار الخط", "trado.ttf")

    font_size = st.slider("حجم الخط الأساسي", 20, 300, 80)

    pdf_dpi = st.selectbox(
        "جودة القالب / دقة الإخراج",
        [150, 200, 300],
        index=1
    )

    output_type = st.radio(
        "نوع الملفات داخل ZIP",
        ["PDF", "JPEG"],
        index=0
    )

    jpeg_quality = 95
    if output_type == "JPEG":
        jpeg_quality = st.slider("جودة JPEG", 70, 100, 95)

    st.markdown("### مكان الاسم")
    name_x = st.slider("X الاسم", 0.0, 1.0, 0.6)
    name_y = st.slider("Y الاسم", 0.0, 1.0, 0.25)
    name_w = st.slider("عرض الاسم", 0.1, 0.9, 0.5)

    st.markdown("### مكان التاريخ")
    date_x = st.slider("X التاريخ", 0.0, 1.0, 0.6)
    date_y = st.slider("Y التاريخ", 0.0, 1.0, 0.3)

    color = st.color_picker("اللون", "#000000")


p_name = Placement(name_x, name_y, name_w)
p_date = Placement(date_x, date_y, 0.4)


if template_file and names_file:
    try:
        template_img = load_template(template_file, dpi=pdf_dpi)
        data = extract_data_from_docx(names_file)

        if data:
            st.success(f"عدد الأسماء: {len(data)}")

            st.markdown("### 👀 معاينة")

            idx = st.selectbox(
                "اختار اسم للتجربة",
                range(len(data)),
                format_func=lambda i: data[i][0]
            )

            name, date = data[idx]

            preview_img = draw_on_template(
                template_img,
                name,
                date,
                font_path,
                font_size,
                p_name,
                p_date,
                color,
                dpi=pdf_dpi
            )

            st.image(preview_img, caption=f"{name} | {date}", width="stretch")

            if output_type == "PDF":
                preview_file = image_to_pdf_bytes(preview_img, dpi=pdf_dpi)
                preview_name = f"{safe_filename(name, 'preview')}.pdf"
                preview_mime = "application/pdf"
                preview_label = "تحميل PDF للمعاينة"
            else:
                preview_file = image_to_jpeg_bytes(preview_img, quality=jpeg_quality)
                preview_name = f"{safe_filename(name, 'preview')}.jpg"
                preview_mime = "image/jpeg"
                preview_label = "تحميل JPEG للمعاينة"

            st.download_button(
                preview_label,
                preview_file,
                file_name=preview_name,
                mime=preview_mime
            )

            if st.button("توليد الشهادات داخل ZIP"):
                zip_buffer = io.BytesIO()
                progress = st.progress(0)

                with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_STORED) as z:
                    for i, (n, d) in enumerate(data, 1):
                        img = draw_on_template(
                            template_img,
                            n,
                            d,
                            font_path,
                            font_size,
                            p_name,
                            p_date,
                            color,
                            dpi=pdf_dpi
                        )

                        safe_name = safe_filename(n, f"cert_{i}")

                        if output_type == "PDF":
                            file_bytes = image_to_pdf_bytes(img, dpi=pdf_dpi)
                            filename = f"{i:03d}_{safe_name}.pdf"
                        else:
                            file_bytes = image_to_jpeg_bytes(img, quality=jpeg_quality)
                            filename = f"{i:03d}_{safe_name}.jpg"

                        z.writestr(filename, file_bytes)

                        del img
                        del file_bytes

                        progress.progress(i / len(data))

                zip_name = "certificates_pdf.zip" if output_type == "PDF" else "certificates_jpeg.zip"

                st.download_button(
                    f"تحميل كل الشهادات {output_type} داخل ZIP",
                    zip_buffer.getvalue(),
                    zip_name,
                    mime="application/zip"
                )

        else:
            st.warning("لم يتم العثور على أسماء داخل ملف الـ Word")

    except Exception as e:
        st.error(f"حصل خطأ: {e}")

else:
    st.info("ارفع الملفات الأول")
