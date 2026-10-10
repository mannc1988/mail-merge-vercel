import io
import os
import zipfile
from pathlib import Path
from typing import Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from docx import Document
import openpyxl

app = FastAPI()

def replace_placeholders_in_doc(doc: Document, replacements: dict):
    """Thay thế {{Key}} trong các Đoạn văn (Paragraphs) và Bảng (Tables)."""
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
    """Hàm gộp nội dung từ source_doc vào target_doc."""
    for element in source_doc.element.body:
        target_doc.element.body.append(element)

@app.get("/", response_class=HTMLResponse)
def read_root():
    """Đọc tệp templates/index.html linh hoạt cho Vercel Serverless."""
    possible_paths = [
        Path(__file__).resolve().parent.parent / "templates" / "index.html",
        Path("templates/index.html"),
        Path(__file__).resolve().parent / "templates" / "index.html"
    ]
    
    for path in possible_paths:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return HTMLResponse(content=f.read())
                
    raise HTTPException(status_code=404, detail="Không tìm thấy tệp templates/index.html")

@app.post("/api/merge")
async def merge_files(
    template_file: UploadFile = File(...),
    data_file: UploadFile = File(...),
    export_mode: str = Form("zip"),
    filename_column: Optional[str] = Form(None)
):
    if not template_file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Tệp mẫu phải có định dạng .docx")
    if not data_file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Tệp dữ liệu phải có định dạng .xlsx")

    try:
        # Đọc dữ liệu Excel
        data_bytes = await data_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)
        sheet = wb.active
        rows = list(sheet.iter_rows(values_only=True))

        if not rows or len(rows) < 2:
            raise HTTPException(status_code=400, detail="File Excel không có dữ liệu!")

        headers = [str(cell).strip() if cell is not None else f"Column_{i}" for i, cell in enumerate(rows[0])]
        data_rows = rows[1:]

        template_bytes = await template_file.read()

        # CHẾ ĐỘ 1: XUẤT 1 FILE WORD DUY NHẤT
        if export_mode == "single_doc":
            merged_doc = Document()
            for idx, row in enumerate(data_rows):
                row_dict = {h: ("" if val is None else str(val)) for h, val in zip(headers, row)}
                
                doc_stream = io.BytesIO(template_bytes)
                doc = Document(doc_stream)
                replace_placeholders_in_doc(doc, row_dict)

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

        # CHẾ ĐỘ 2: XUẤT TỆP ZIP CHỨA NHIỀU FILE WORD
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, row in enumerate(data_rows, start=1):
                    row_dict = {h: ("" if val is None else str(val)) for h, val in zip(headers, row)}

                    doc_stream = io.BytesIO(template_bytes)
                    doc = Document(doc_stream)
                    replace_placeholders_in_doc(doc, row_dict)

                    out_filename = f"Document_{idx}.docx"
                    if filename_column and filename_column in row_dict and row_dict[filename_column]:
                        safe_name = "".join([c for c in row_dict[filename_column] if c.isalnum() or c in (" ", "_", "-")]).strip()
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