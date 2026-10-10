import io
import os
import zipfile
from datetime import datetime, date
from pathlib import Path
from typing import Optional, Any
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
import openpyxl

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def format_cell_value(val: Any, num_format_type: str = "VN") -> str:
    """Xử lý định dạng theo 3 kiểu dữ liệu: Date, Number (VN/US), String."""
    if val is None:
        return ""

    if isinstance(val, (datetime, date)):
        return val.strftime("%d/%m/%Y")

    is_num = False
    num_val = None

    if isinstance(val, (int, float)):
        is_num = True
        num_val = val
    elif isinstance(val, str):
        cleaned_str = val.strip()
        temp_str = cleaned_str.replace(',', '.')
        try:
            num_val = float(temp_str) if '.' in temp_str else int(temp_str)
            is_num = True
        except ValueError:
            is_num = False

    if is_num and num_val is not None:
        fmt_type = str(num_format_type).strip().upper()
        
        if isinstance(num_val, float):
            str_float = f"{num_val:.10f}".rstrip('0').rstrip('.')
            if '.' in str_float:
                int_part, dec_part = str_float.split('.')
                formatted_int = f"{int(int_part):,}"
                formatted = f"{formatted_int}.{dec_part}"
            else:
                formatted = f"{int(num_val):,}"
        else:
            formatted = f"{int(num_val):,}"

        if fmt_type in ["VN", ".", "VIETNAM", "VIỆT NAM"]:
            formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
            
        return formatted

    return str(val)

def replace_placeholders_in_doc(doc: Document, replacements: dict):
    """Thay thế {{Key}} trong Đoạn văn và Bảng."""
    for p in doc.paragraphs:
        for key, val in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(val))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(val))

def append_document(target_doc: Document, source_doc: Document):
    """Gộp nội dung từ source_doc vào target_doc."""
    for element in source_doc.element.body:
        target_doc.element.body.append(element)

def get_template_bytes(template_name: Optional[str] = None, uploaded_bytes: Optional[bytes] = None) -> bytes:
    """Lấy file mẫu: Ưu tiên file upload custom từ client ➔ tệp theo cột BieuMau ➔ template.docx mặc định."""
    if uploaded_bytes:
        return uploaded_bytes

    filename = str(template_name).strip() if template_name and str(template_name).strip() else "template.docx"
    if not filename.endswith(".docx"):
        filename += ".docx"

    possible_paths = [
        Path(__file__).resolve().parent.parent / "templates" / filename,
        Path(__file__).resolve().parent.parent / filename,
        Path(__file__).resolve().parent / filename,
        Path("templates") / filename,
        Path(filename)
    ]
    for path in possible_paths:
        if path.exists():
            with open(path, "rb") as f:
                return f.read()

    default_paths = [
        Path(__file__).resolve().parent.parent / "template.docx",
        Path("template.docx")
    ]
    for path in default_paths:
        if path.exists():
            with open(path, "rb") as f:
                return f.read()

    raise HTTPException(status_code=500, detail=f"Không tìm thấy biểu mẫu '{filename}' trên Server!")

@app.get("/", response_class=HTMLResponse)
def read_root():
    possible_paths = [
        Path(__file__).resolve().parent.parent / "index.html",
        Path("index.html"),
        Path(__file__).resolve().parent.parent / "templates" / "index.html",
        Path("templates/index.html")
    ]
    for path in possible_paths:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return HTMLResponse(content=f.read())
    raise HTTPException(status_code=404, detail="Không tìm thấy tệp index.html")

@app.post("/api/merge")
@app.post("/merge")
async def merge_files(
    data_file: UploadFile = File(...),
    template_file: Optional[UploadFile] = File(None), # 💡 Hỗ trợ upload file mẫu custom
    export_mode: str = Form("zip"),
    filename_column: Optional[str] = Form(None)
):
    if not data_file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Tệp dữ liệu phải có định dạng .xlsx")

    # Đọc bytes của tệp template tùy chọn gửi từ client
    custom_template_bytes = None
    if template_file and template_file.filename.endswith(".docx"):
        custom_template_bytes = await template_file.read()

    try:
        data_bytes = await data_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)
        sheet = wb.active
        rows = list(sheet.iter_rows(values_only=True))

        if not rows or len(rows) < 2:
            raise HTTPException(status_code=400, detail="File Excel không có dữ liệu!")

        headers = [str(cell).strip() if cell is not None else f"Column_{i}" for i, cell in enumerate(rows[0])]
        data_rows = rows[1:]

        # Lọc danh sách dữ liệu hợp lệ (xử lý TrangThai)
        valid_rows = []
        for row in data_rows:
            row_dict = {h: val for h, val in zip(headers, row)}
            
            status_val = str(row_dict.get("TrangThai", "") or "").strip().lower()
            if status_val in ["ân", "an", "bỏ qua", "bo qua", "ignore", "0", "false", "draft"]:
                continue
            
            valid_rows.append(row_dict)

        if not valid_rows:
            raise HTTPException(status_code=400, detail="Không có dòng dữ liệu hợp lệ để xử lý!")

        # CHẾ ĐỘ 1: XUẤT 1 FILE WORD DUY NHẤT
        if export_mode == "single_doc":
            merged_doc = Document()
            for idx, row_dict in enumerate(valid_rows):
                num_fmt = str(row_dict.get("DauPhanCachPhanNghin", "VN") or "VN").strip().upper()

                replacements = {}
                for k, v in row_dict.items():
                    replacements[k] = format_cell_value(v, num_fmt)

                template_name = str(row_dict.get("BieuMau", "") or "").strip()
                template_bytes = get_template_bytes(template_name, custom_template_bytes)

                doc_stream = io.BytesIO(template_bytes)
                doc = Document(doc_stream)
                replace_placeholders_in_doc(doc, replacements)

                if idx > 0:
                    merged_doc.add_page_break()

                append_document(merged_doc, doc)

            output_stream = io.BytesIO()
            merged_doc.save(output_stream)
            output_stream.seek(0)

            return Response(
                content=output_stream.getvalue(),
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": "attachment; filename=Mail_Merge_Merged.docx"}
            )

        # CHẾ ĐỘ 2: XUẤT TỆP ZIP
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, row_dict in enumerate(valid_rows, start=1):
                    num_fmt = str(row_dict.get("DauPhanCachPhanNghin", "VN") or "VN").strip().upper()

                    replacements = {}
                    for k, v in row_dict.items():
                        replacements[k] = format_cell_value(v, num_fmt)

                    template_name = str(row_dict.get("BieuMau", "") or "").strip()
                    template_bytes = get_template_bytes(template_name, custom_template_bytes)

                    doc_stream = io.BytesIO(template_bytes)
                    doc = Document(doc_stream)
                    replace_placeholders_in_doc(doc, replacements)

                    out_filename = f"Document_{idx}.docx"
                    if filename_column and filename_column in replacements and replacements[filename_column]:
                        safe_name = "".join([c for c in str(replacements[filename_column]) if c.isalnum() or c in (" ", "_", "-")]).strip()
                        if safe_name:
                            out_filename = f"{safe_name}_{idx}.docx"

                    output_doc_stream = io.BytesIO()
                    doc.save(output_doc_stream)
                    output_doc_stream.seek(0)
                    zip_file.writestr(out_filename, output_doc_stream.getvalue())

            zip_buffer.seek(0)
            return Response(
                content=zip_buffer.getvalue(),
                media_type="application/zip",
                headers={"Content-Disposition": "attachment; filename=Mail_Merge_Results.zip"}
            )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý: {str(e)}")