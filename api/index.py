import io
import os
import mimetypes
import base64
import zipfile
import requests
from datetime import datetime, date
from pathlib import Path
from typing import Optional, Any
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from spire.doc import Document as SpireDocument, FileFormat as DocFileFormat
from spire.xls import Workbook as SpireWorkbook, FileFormat as XlsFileFormat
import openpyxl

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

GITHUB_REPO = "mannc1988/mail-merge-vercel"
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")


def create_default_template_docx_bytes() -> bytes:
    """Tự động tạo tệp Word mẫu nếu thiếu file trên server."""
    doc = SpireDocument()
    section = doc.AddSection()
    
    p_title = section.AddParagraph()
    p_title.AppendText("GIẤY XÁC NHẬN / THÔNG BÁO")
    
    p1 = section.AddParagraph()
    p1.AppendText("Kính gửi ông/bà: {{HoTen}}")
    
    p2 = section.AddParagraph()
    p2.AppendText("Ngày sinh: {{NgaySinh}}\nChức vụ: {{ChucVu}}\nSố tiền thanh toán: {{SoTien}} VNĐ")
    
    p3 = section.AddParagraph()
    p3.AppendText("\nTrân trọng cảm ơn!")
    
    stream = io.BytesIO()
    doc.SaveToStream(stream, DocFileFormat.Docx)
    doc.Close()
    stream.seek(0)
    return stream.getvalue()


def format_cell_value(val: Any, num_format_type: str = "VN") -> str:
    """Định dạng dữ liệu số tiền, ngày tháng."""
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


def replace_placeholders_spire(doc: SpireDocument, replacements: dict):
    """Thay thế {{Key}} trong Word bằng Spire.Doc Engine."""
    for key, val in replacements.items():
        placeholder = f"{{{{{key}}}}}"
        doc.Replace(placeholder, str(val), False, True)


def get_template_bytes(template_name: Optional[str] = None, custom_bytes: Optional[bytes] = None) -> bytes:
    if custom_bytes:
        return custom_bytes

    filename = str(template_name).strip() if template_name and str(template_name).strip() else "template.docx"
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

    return create_default_template_docx_bytes()


@app.get("/", response_class=HTMLResponse)
def read_root():
    possible_paths = [
        Path(__file__).resolve().parent.parent / "index.html",
        Path("index.html")
    ]
    for path in possible_paths:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return HTMLResponse(content=f.read())
    raise HTTPException(status_code=404, detail="Không tìm thấy tệp index.html")


@app.get("/api/list-templates")
def list_templates():
    templates = set()
    templates_dir = Path(__file__).resolve().parent.parent / "templates"
    if templates_dir.exists():
        for f in templates_dir.glob("*.*"):
            templates.add(f.name)
            
    root_dir = Path(__file__).resolve().parent.parent
    if root_dir.exists():
        for f in root_dir.glob("*.*"):
            if f.suffix.lower() in [".docx", ".pdf", ".xlsx", ".png", ".jpg", ".jpeg"]:
                templates.add(f.name)

    if not templates:
        templates.add("template.docx")

    return {"templates": sorted(list(templates))}


# 🎯 API XEM TỆP TỪ SERVER: RENDER TRỰC TIẾP QUA SPIRE (OFFICE) HOẶC NATIVE (PDF/ẢNH)
@app.get("/api/preview-spire-remote")
def preview_spire_remote(name: str):
    file_bytes = get_template_bytes(template_name=name)
    ext = Path(name).suffix.lower()

    # 1. Word -> Render trực tiếp HTML qua Spire.Doc Engine
    if ext in [".docx", ".doc"]:
        doc = SpireDocument()
        doc.LoadFromStream(io.BytesIO(file_bytes), DocFileFormat.Docx)
        html_stream = io.BytesIO()
        doc.SaveToStream(html_stream, DocFileFormat.Html)
        doc.Close()
        html_stream.seek(0)
        return HTMLResponse(content=html_stream.getvalue().decode('utf-8', errors='ignore'))

    # 2. Excel -> Render trực tiếp HTML qua Spire.XLS Engine
    elif ext in [".xlsx", ".xls"]:
        workbook = SpireWorkbook()
        workbook.LoadFromStream(io.BytesIO(file_bytes))
        html_stream = io.BytesIO()
        workbook.SaveToStream(html_stream, XlsFileFormat.HTML)
        workbook.Dispose()
        html_stream.seek(0)
        return HTMLResponse(content=html_stream.getvalue().decode('utf-8', errors='ignore'))

    # 3. PDF & Ảnh -> Trả về Media gốc
    mime_type, _ = mimetypes.guess_type(name)
    return Response(content=file_bytes, media_type=mime_type or "application/octet-stream")


# 🎯 API XEM TỆP LOCAL TẢI LÊN: RENDER TRỰC TIẾP QUA SPIRE
@app.post("/api/preview-spire-local")
async def preview_spire_local(file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()
    file_bytes = await file.read()

    if ext in [".docx", ".doc"]:
        doc = SpireDocument()
        doc.LoadFromStream(io.BytesIO(file_bytes), DocFileFormat.Docx)
        html_stream = io.BytesIO()
        doc.SaveToStream(html_stream, DocFileFormat.Html)
        doc.Close()
        html_stream.seek(0)
        return HTMLResponse(content=html_stream.getvalue().decode('utf-8', errors='ignore'))

    elif ext in [".xlsx", ".xls"]:
        workbook = SpireWorkbook()
        workbook.LoadFromStream(io.BytesIO(file_bytes))
        html_stream = io.BytesIO()
        workbook.SaveToStream(html_stream, XlsFileFormat.HTML)
        workbook.Dispose()
        html_stream.seek(0)
        return HTMLResponse(content=html_stream.getvalue().decode('utf-8', errors='ignore'))

    mime_type = file.content_type or mimetypes.guess_type(file.filename)[0]
    return Response(content=file_bytes, media_type=mime_type or "application/octet-stream")


@app.get("/api/get-template-file")
def get_template_file(name: str):
    template_bytes = get_template_bytes(template_name=name)
    return Response(
        content=template_bytes,
        media_type="application/octet-stream",
        headers={"Content-Disposition": f"attachment; filename={name}"}
    )


@app.get("/api/download-sample/template")
def download_sample_template():
    template_bytes = get_template_bytes()
    return Response(
        content=template_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=Mau_Word_Template.docx"}
    )


@app.get("/api/download-sample/excel")
def download_sample_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DuLieuMau"

    ws.append(["HoTen", "NgaySinh", "ChucVu", "SoTien", "BieuMau", "TrangThai", "DauPhanCachPhanNghin"])
    ws.append(["Nguyễn Văn A", "2000-05-15", "Trưởng phòng", 15000000.75, "template.docx", "Hiện", "VN"])
    ws.append(["Trần Thị B", "1995-10-20", "Chuyên viên", 8500000, "template.docx", "Hiện", "US"])
    ws.append(["Lê Văn C", "1998-12-01", "Nhân viên", 5000000, "template.docx", "Ẩn", "VN"])

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    return Response(
        content=output_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Mau_Excel_Data.xlsx"}
    )


@app.post("/api/upload-template")
async def upload_template_to_github(template_file: UploadFile = File(...)):
    if not GITHUB_TOKEN:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GITHUB_TOKEN!")

    filename = template_file.filename.strip()
    file_bytes = await template_file.read()
    content_b64 = base64.b64encode(file_bytes).decode("utf-8")

    github_file_path = f"templates/{filename}"
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{github_file_path}"

    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    sha = None
    get_res = requests.get(url, headers=headers)
    if get_res.status_code == 200:
        sha = get_res.json().get("sha")

    payload = {
        "message": f"Upload new template {filename} via Web Interface",
        "content": content_b64,
        "branch": "main"
    }
    if sha:
        payload["sha"] = sha

    put_res = requests.put(url, json=payload, headers=headers)

    if put_res.status_code in [200, 201]:
        return {"message": f"Đã lưu thành công biểu mẫu '{filename}' lên GitHub!", "filename": filename}
    else:
        err_msg = put_res.json().get("message", "Lỗi không xác định")
        raise HTTPException(status_code=500, detail=f"Không thể lưu biểu mẫu lên GitHub: {err_msg}")


@app.post("/api/merge")
@app.post("/merge")
async def merge_files(
    data_file: UploadFile = File(...),
    template_file: Optional[UploadFile] = File(None),
    export_mode: str = Form("zip"),
    filename_column: Optional[str] = Form(None)
):
    if not data_file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Tệp dữ liệu phải có định dạng .xlsx")

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

        valid_rows = []
        for row in data_rows:
            row_dict = {h: val for h, val in zip(headers, row)}
            
            status_val = str(row_dict.get("TrangThai", "") or "").strip().lower()
            if status_val in ["ân", "an", "bỏ qua", "bo qua", "ignore", "0", "false", "draft"]:
                continue
            
            valid_rows.append(row_dict)

        if not valid_rows:
            raise HTTPException(status_code=400, detail="Không có dòng dữ liệu hợp lệ để xử lý!")

        if export_mode == "single_doc":
            merged_doc = SpireDocument()
            
            for idx, row_dict in enumerate(valid_rows):
                num_fmt = str(row_dict.get("DauPhanCachPhanNghin", "VN") or "VN").strip().upper()

                replacements = {}
                for k, v in row_dict.items():
                    replacements[k] = format_cell_value(v, num_fmt)

                template_name = str(row_dict.get("BieuMau", "") or "").strip()
                template_bytes = get_template_bytes(template_name, custom_template_bytes)

                doc_temp = SpireDocument()
                doc_temp.LoadFromStream(io.BytesIO(template_bytes), DocFileFormat.Docx)
                
                replace_placeholders_spire(doc_temp, replacements)

                for i in range(doc_temp.Sections.Count):
                    sec = doc_temp.Sections.get_Item(i).Clone()
                    merged_doc.Sections.Add(sec)

                doc_temp.Close()

            output_stream = io.BytesIO()
            merged_doc.SaveToStream(output_stream, DocFileFormat.Docx)
            merged_doc.Close()
            output_stream.seek(0)

            return Response(
                content=output_stream.getvalue(),
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": "attachment; filename=Mail_Merge_Merged_Spire.docx"}
            )

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

                    doc = SpireDocument()
                    doc.LoadFromStream(io.BytesIO(template_bytes), DocFileFormat.Docx)
                    
                    replace_placeholders_spire(doc, replacements)

                    out_filename = f"Document_{idx}.docx"
                    if filename_column and filename_column in replacements and replacements[filename_column]:
                        safe_name = "".join([c for c in str(replacements[filename_column]) if c.isalnum() or c in (" ", "_", "-")]).strip()
                        if safe_name:
                            out_filename = f"{safe_name}_{idx}.docx"

                    output_doc_stream = io.BytesIO()
                    doc.SaveToStream(output_doc_stream, DocFileFormat.Docx)
                    doc.Close()
                    output_doc_stream.seek(0)
                    
                    zip_file.writestr(out_filename, output_doc_stream.getvalue())

            zip_buffer.seek(0)
            return Response(
                content=zip_buffer.getvalue(),
                media_type="application/zip",
                headers={"Content-Disposition": "attachment; filename=Mail_Merge_Results_Spire.zip"}
            )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý Spire: {str(e)}")