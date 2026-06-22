# File → Markdown App

Ứng dụng web nội bộ giúp chuyển đổi **nhiều file cùng lúc** sang Markdown (`.md`).

## 1. Tính năng

- Kéo-thả nhiều file trong một lần.
- Hỗ trợ:
  - Word: `.docx`
  - Excel: `.xlsx`, `.xlsm`
  - PowerPoint: `.pptx`
  - PDF: `.pdf`
  - Ảnh: `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tiff`
  - Văn bản: `.txt`, `.md`, `.csv`
- Giữ cấu trúc cơ bản:
  - Tiêu đề và đoạn văn Word
  - Bảng Word/Excel
  - Nội dung theo từng slide PowerPoint
  - Nội dung theo từng trang PDF
- OCR cho ảnh và PDF scan.
- Xem trước nội dung Markdown.
- Tải từng file `.md` hoặc tải toàn bộ thành `.zip`.
- Chạy cục bộ: tài liệu không tự gửi ra dịch vụ bên ngoài.

## 2. Cài đặt trên Windows

### Bước 1 — Cài Python

Cài Python 3.11 hoặc 3.12 từ trang chính thức của Python. Khi cài, chọn:

`Add Python to PATH`

### Bước 2 — Mở thư mục ứng dụng

Mở PowerShell trong thư mục này.

### Bước 3 — Tạo môi trường ảo

```powershell
python -m venv .venv
.venv\Scripts\activate
```

### Bước 4 — Cài thư viện

```powershell
pip install -r requirements.txt
```

### Bước 5 — Cài Tesseract OCR

OCR chỉ cần khi muốn đọc ảnh hoặc PDF scan.

1. Cài Tesseract OCR cho Windows.
2. Bảo đảm `tesseract.exe` nằm trong PATH.
3. Cài thêm dữ liệu ngôn ngữ:
   - `vie.traineddata`
   - `eng.traineddata`

Nếu Tesseract không nằm trong PATH, thêm dòng sau vào đầu file `app.py`
sau phần import:

```python
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

### Bước 6 — Chạy ứng dụng

```powershell
streamlit run app.py
```

Trình duyệt sẽ mở tại địa chỉ thường là:

`http://localhost:8501`

## 3. Chạy nhanh

Có thể nhấp đúp file:

`run_app.bat`

sau khi đã cài Python và các thư viện.

## 4. Lưu ý về chất lượng chuyển đổi

- `.docx`, `.xlsx`, `.pptx` dạng văn bản sẽ cho kết quả tốt nhất.
- PDF có text thật sẽ trích xuất tốt hơn PDF scan.
- OCR phụ thuộc chất lượng ảnh, độ nghiêng, độ phân giải và font.
- Biểu đồ, sơ đồ SmartArt, công thức phức tạp và bố cục nhiều cột có thể
  không chuyển hoàn toàn sang Markdown.
- File `.doc`, `.xls`, `.ppt` đời cũ chưa được hỗ trợ trực tiếp; hãy mở bằng
  Microsoft Office và lưu lại thành `.docx`, `.xlsx`, `.pptx`.

## 5. Cấu trúc đầu ra

Mỗi file được chuyển thành một file Markdown riêng:

```markdown
---
title: "Tên tài liệu"
source_file: "Tên tài liệu.docx"
source_type: "docx"
---

# Tên tài liệu

...
```
