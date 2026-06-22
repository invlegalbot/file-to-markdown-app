
import io
import os
import re
import zipfile
import tempfile
from pathlib import Path
from typing import List, Tuple

import streamlit as st
from PIL import Image

# Optional imports are handled gracefully.
try:
    from docx import Document
except Exception:
    Document = None

try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    from pptx import Presentation
except Exception:
    Presentation = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    import pytesseract
except Exception:
    pytesseract = None


SUPPORTED_EXTENSIONS = {
    ".docx", ".xlsx", ".xlsm", ".pptx", ".pdf",
    ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif",
    ".txt", ".md", ".csv"
}


def clean_text(text: str) -> str:
    """Normalize extracted text while preserving paragraph boundaries."""
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def escape_md_cell(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", "<br>").replace("|", r"\|")
    return text.strip()


def table_to_markdown(rows: List[List[str]]) -> str:
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    normalized = [r + [""] * (width - len(r)) for r in rows]
    header = normalized[0]
    body = normalized[1:]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for row in body:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def make_frontmatter(filename: str, source_type: str) -> str:
    safe_name = filename.replace('"', "'")
    return (
        "---\n"
        f'title: "{Path(safe_name).stem}"\n'
        f'source_file: "{safe_name}"\n'
        f'source_type: "{source_type}"\n'
        "---\n\n"
    )


def convert_docx(data: bytes, filename: str) -> str:
    if Document is None:
        raise RuntimeError("Thiếu thư viện python-docx.")
    doc = Document(io.BytesIO(data))
    parts = [make_frontmatter(filename, "docx"), f"# {Path(filename).stem}\n"]

    for block in doc.element.body.iterchildren():
        tag = block.tag.split("}")[-1]

        if tag == "p":
            # Resolve paragraph object from XML element
            para = next((p for p in doc.paragraphs if p._p is block), None)
            if para is None:
                continue
            text = clean_text(para.text)
            if not text:
                continue

            style = (para.style.name or "").lower() if para.style else ""
            if "heading 1" in style or "tiêu đề 1" in style:
                parts.append(f"# {text}")
            elif "heading 2" in style or "tiêu đề 2" in style:
                parts.append(f"## {text}")
            elif "heading 3" in style or "tiêu đề 3" in style:
                parts.append(f"### {text}")
            elif "heading 4" in style or "tiêu đề 4" in style:
                parts.append(f"#### {text}")
            elif "list" in style:
                parts.append(f"- {text}")
            else:
                parts.append(text)

        elif tag == "tbl":
            table = next((t for t in doc.tables if t._tbl is block), None)
            if table is None:
                continue
            rows = []
            for row in table.rows:
                rows.append([escape_md_cell(cell.text) for cell in row.cells])
            md_table = table_to_markdown(rows)
            if md_table:
                parts.append(md_table)

    return clean_text("\n\n".join(parts)) + "\n"


def convert_xlsx(data: bytes, filename: str) -> str:
    if openpyxl is None:
        raise RuntimeError("Thiếu thư viện openpyxl.")
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False, read_only=True)
    parts = [make_frontmatter(filename, "excel"), f"# {Path(filename).stem}\n"]

    for ws in wb.worksheets:
        parts.append(f"## Sheet: {ws.title}")
        rows = []
        for row in ws.iter_rows(values_only=True):
            values = [escape_md_cell(v) for v in row]
            if any(v != "" for v in values):
                rows.append(values)

        if not rows:
            parts.append("_Sheet không có dữ liệu._")
            continue

        # Trim trailing empty columns
        max_non_empty = 0
        for row in rows:
            for idx, value in enumerate(row, start=1):
                if value != "":
                    max_non_empty = max(max_non_empty, idx)
        rows = [row[:max_non_empty] for row in rows]

        # Split large sheets to keep Markdown usable.
        chunk_size = 200
        for start in range(0, len(rows), chunk_size):
            chunk = rows[start:start + chunk_size]
            if start > 0:
                parts.append(f"### Dòng {start + 1}–{start + len(chunk)}")
            parts.append(table_to_markdown(chunk))

    return clean_text("\n\n".join(parts)) + "\n"


def extract_shape_text(shape) -> List[str]:
    items = []
    if hasattr(shape, "text") and shape.text:
        text = clean_text(shape.text)
        if text:
            items.append(text)

    if getattr(shape, "has_table", False):
        rows = []
        for row in shape.table.rows:
            rows.append([escape_md_cell(cell.text) for cell in row.cells])
        md = table_to_markdown(rows)
        if md:
            items.append(md)

    if getattr(shape, "shape_type", None) == 6 and hasattr(shape, "shapes"):  # group
        for child in shape.shapes:
            items.extend(extract_shape_text(child))

    return items


def convert_pptx(data: bytes, filename: str) -> str:
    if Presentation is None:
        raise RuntimeError("Thiếu thư viện python-pptx.")
    prs = Presentation(io.BytesIO(data))
    parts = [make_frontmatter(filename, "pptx"), f"# {Path(filename).stem}\n"]

    for idx, slide in enumerate(prs.slides, start=1):
        title = ""
        if slide.shapes.title and slide.shapes.title.text:
            title = clean_text(slide.shapes.title.text)
        heading = f"## Slide {idx}" + (f": {title}" if title else "")
        parts.append(heading)

        slide_items = []
        for shape in slide.shapes:
            if slide.shapes.title is not None and shape == slide.shapes.title:
                continue
            slide_items.extend(extract_shape_text(shape))

        if slide_items:
            for item in slide_items:
                if "\n" in item and not item.startswith("|"):
                    for line in item.splitlines():
                        line = line.strip()
                        if line:
                            slide_items_line = f"- {line}"
                            parts.append(slide_items_line)
                else:
                    parts.append(item)
        else:
            parts.append("_Không phát hiện nội dung văn bản trên slide._")

        # Speaker notes, when available
        try:
            notes = slide.notes_slide.notes_text_frame.text
            notes = clean_text(notes)
            if notes:
                parts.append("### Ghi chú diễn giả")
                parts.append(notes)
        except Exception:
            pass

    return clean_text("\n\n".join(parts)) + "\n"


def ocr_image(image: Image.Image, lang: str) -> str:
    if pytesseract is None:
        raise RuntimeError("OCR chưa sẵn sàng: thiếu pytesseract.")
    try:
        return clean_text(pytesseract.image_to_string(image, lang=lang))
    except pytesseract.TesseractNotFoundError:
        raise RuntimeError(
            "Chưa cài Tesseract OCR trên máy. Xem hướng dẫn trong README."
        )


def convert_image(data: bytes, filename: str, use_ocr: bool, ocr_lang: str) -> str:
    image = Image.open(io.BytesIO(data))
    parts = [make_frontmatter(filename, "image"), f"# {Path(filename).stem}\n"]
    parts.append(f"- Kích thước ảnh: {image.width} × {image.height} px")
    parts.append(f"- Định dạng: {image.format or Path(filename).suffix.lstrip('.')}")
    if use_ocr:
        text = ocr_image(image, ocr_lang)
        parts.append("\n## Nội dung OCR")
        parts.append(text if text else "_Không nhận diện được văn bản._")
    else:
        parts.append("\n_Đã tắt OCR. Bật tùy chọn OCR để nhận diện chữ trong ảnh._")
    return clean_text("\n\n".join(parts)) + "\n"


def convert_pdf(data: bytes, filename: str, use_ocr: bool, ocr_lang: str) -> str:
    if PdfReader is None:
        raise RuntimeError("Thiếu thư viện pypdf.")
    parts = [make_frontmatter(filename, "pdf"), f"# {Path(filename).stem}\n"]
    reader = PdfReader(io.BytesIO(data))

    extracted_any = False
    page_texts = []
    for idx, page in enumerate(reader.pages, start=1):
        text = clean_text(page.extract_text() or "")
        if text:
            extracted_any = True
        page_texts.append(text)

    for idx, text in enumerate(page_texts, start=1):
        parts.append(f"## Trang {idx}")
        if text:
            parts.append(text)
        elif use_ocr:
            if fitz is None:
                parts.append("_Không có text; thiếu PyMuPDF để OCR trang PDF._")
                continue
            try:
                pdf_doc = fitz.open(stream=data, filetype="pdf")
                page = pdf_doc.load_page(idx - 1)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image = Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_text = ocr_image(image, ocr_lang)
                parts.append(ocr_text if ocr_text else "_Không nhận diện được văn bản._")
            except Exception as exc:
                parts.append(f"_OCR thất bại: {exc}_")
        else:
            parts.append("_Không trích xuất được văn bản. Có thể đây là trang scan; hãy bật OCR._")

    return clean_text("\n\n".join(parts)) + "\n"


def convert_plain_text(data: bytes, filename: str) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1258", "latin-1"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = data.decode("utf-8", errors="replace")

    if Path(filename).suffix.lower() == ".md":
        return text
    return make_frontmatter(filename, "text") + f"# {Path(filename).stem}\n\n" + clean_text(text) + "\n"


def convert_file(data: bytes, filename: str, use_ocr: bool, ocr_lang: str) -> str:
    ext = Path(filename).suffix.lower()
    if ext == ".docx":
        return convert_docx(data, filename)
    if ext in {".xlsx", ".xlsm"}:
        return convert_xlsx(data, filename)
    if ext == ".pptx":
        return convert_pptx(data, filename)
    if ext == ".pdf":
        return convert_pdf(data, filename, use_ocr, ocr_lang)
    if ext in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif"}:
        return convert_image(data, filename, use_ocr, ocr_lang)
    if ext in {".txt", ".md", ".csv"}:
        return convert_plain_text(data, filename)
    raise ValueError(f"Định dạng chưa được hỗ trợ: {ext}")


def zip_outputs(outputs: List[Tuple[str, str]]) -> bytes:
    buffer = io.BytesIO()
    used_names = set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for original_name, markdown in outputs:
            base = Path(original_name).stem
            name = f"{base}.md"
            counter = 2
            while name.lower() in used_names:
                name = f"{base}_{counter}.md"
                counter += 1
            used_names.add(name.lower())
            zf.writestr(name, markdown.encode("utf-8"))
    return buffer.getvalue()


st.set_page_config(
    page_title="File → Markdown",
    page_icon="📝",
    layout="wide",
)

st.title("📝 Chuyển đổi nhiều file sang Markdown")
st.caption(
    "Kéo-thả Word, Excel, PDF, PowerPoint hoặc ảnh. "
    "Ứng dụng xử lý trên máy đang chạy app và không tự gửi tài liệu ra ngoài."
)

with st.sidebar:
    st.header("Thiết lập")
    use_ocr = st.toggle("Bật OCR cho ảnh/PDF scan", value=True)
    ocr_lang = st.selectbox(
        "Ngôn ngữ OCR",
        options=["vie+eng", "vie", "eng"],
        index=0,
        help="Cần cài bộ ngôn ngữ tương ứng cho Tesseract OCR."
    )
    st.markdown(
        """
**Định dạng hỗ trợ**

- Word: `.docx`
- Excel: `.xlsx`, `.xlsm`
- PowerPoint: `.pptx`
- PDF: `.pdf`
- Ảnh: `.png`, `.jpg`, `.webp`, `.bmp`, `.tiff`
- Văn bản: `.txt`, `.md`, `.csv`
        """
    )

uploaded_files = st.file_uploader(
    "Chọn hoặc kéo-thả nhiều file",
    type=[e.lstrip(".") for e in sorted(SUPPORTED_EXTENSIONS)],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"Đã chọn **{len(uploaded_files)}** file.")

    if st.button("Chuyển đổi sang Markdown", type="primary", use_container_width=True):
        outputs = []
        failures = []

        progress = st.progress(0, text="Bắt đầu chuyển đổi...")
        for idx, uploaded in enumerate(uploaded_files, start=1):
            try:
                markdown = convert_file(
                    uploaded.getvalue(),
                    uploaded.name,
                    use_ocr=use_ocr,
                    ocr_lang=ocr_lang,
                )
                outputs.append((uploaded.name, markdown))
            except Exception as exc:
                failures.append((uploaded.name, str(exc)))

            progress.progress(
                idx / len(uploaded_files),
                text=f"Đang xử lý {idx}/{len(uploaded_files)}: {uploaded.name}"
            )

        progress.empty()
        st.session_state["outputs"] = outputs
        st.session_state["failures"] = failures

outputs = st.session_state.get("outputs", [])
failures = st.session_state.get("failures", [])

if failures:
    st.error(f"Có {len(failures)} file chuyển đổi chưa thành công.")
    for name, error in failures:
        st.write(f"**{name}:** {error}")

if outputs:
    st.success(f"Đã chuyển đổi thành công **{len(outputs)}** file.")

    zip_bytes = zip_outputs(outputs)
    st.download_button(
        "⬇️ Tải toàn bộ file Markdown (.zip)",
        data=zip_bytes,
        file_name="markdown_outputs.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )

    st.divider()
    for original_name, markdown in outputs:
        md_name = f"{Path(original_name).stem}.md"
        with st.expander(f"📄 {original_name} → {md_name}", expanded=False):
            st.download_button(
                f"Tải {md_name}",
                data=markdown.encode("utf-8"),
                file_name=md_name,
                mime="text/markdown",
                key=f"download_{original_name}_{len(markdown)}",
            )
            st.text_area(
                "Xem trước / chỉnh sửa nội dung",
                value=markdown,
                height=420,
                key=f"preview_{original_name}_{len(markdown)}",
            )
