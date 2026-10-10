import io
import os
import zipfile
from typing import Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, Response
from docx import Document
import openpyxl

app = FastAPI()

# Giao diện HTML + Bootstrap 5 + DataTables + FontAwesome nhúng trực tiếp
HTML_CONTENT = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Công cụ Mail Merge Trộn Biểu Mẫu Word & Excel</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/1.13.7/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f6f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .card-custom { border: none; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
        .upload-box { border: 2px dashed #cbd5e1; border-radius: 8px; padding: 20px; text-align: center; background-color: #f8fafc; transition: all 0.3s; }
        .upload-box:hover { border-color: #0d6efd; background-color: #eff6ff; }
    </style>
</head>
<body>

<div class="container py-5">
    <div class="text-center mb-4">
        <h2 class="fw-bold text-primary"><i class="fa-solid fa-file-word me-2"></i>Công Cụ Mail Merge Online</h2>
        <p class="text-muted">Tự động trộn dữ liệu từ Excel vào tệp mẫu Word (<code class="bg-light px-2 py-1 rounded text-danger">{{TenBiên}}</code>)</p>
    </div>

    <div class="card card-custom p-4 mb-4">
        <form id="mergeForm">
            <div class="row g-4 mb-4">
                <div class="col-md-6">
                    <label class="form-label fw-semibold"><i class="fa-solid fa-file-word text-primary me-2"></i>1. Tệp Word Mẫu (.docx)</label>
                    <div class="upload-box">
                        <i class="fa-solid fa-cloud-arrow-up fa-2x text-primary mb-2"></i>
                        <input type="file" id="templateFile" accept=".docx" class="form-control" required />
                        <small class="text-muted d-block mt-1">Chứa các nhãn dạng {{Ten}}, {{ChucVu}}...</small>
                    </div>
                </div>

                <div class="col-md-6">
                    <label class="form-label fw-semibold"><i class="fa-solid fa-file-excel text-success me-2"></i>2. Tệp Excel Dữ Liệu (.xlsx)</label>
                    <div class="upload-box">
                        <i class="fa-solid fa-file-csv fa-2x text-success mb-2"></i>
                        <input type="file" id="dataFile" accept=".xlsx" class="form-control" required />
                        <small class="text-muted d-block mt-1">Hàng đầu tiên là tên cột tương ứng</small>
                    </div>
                </div>
            </div>

            <div class="row g-3 mb-4 bg-light p-3 rounded">
                <div class="col-md-6">
                    <label class="form-label fw-semibold"><i class="fa-solid fa-signature me-2"></i>Cột Excel dùng đặt tên file:</label>
                    <input type="text" id="filenameColumn" placeholder="Ví dụ: HoTen hoặc MaSV (Tùy chọn)" class="form-control" />
                </div>
                <div class="col-md-6">
                    <label class="form-label fw-semibold"><i class="fa-solid fa-file-export me-2"></i>Định dạng xuất kết quả:</label>
                    <select id="exportMode" class="form-select">
                        <option value="zip">Xuất từng file riêng lẻ đóng gói ZIP</option>
                        <option value="single_doc">Trộn tất cả gộp vào 1 file Word duy nhất</option>
                    </select>
                </div>
            </div>

            <div class="text-center">
                <button type="submit" id="submitBtn" class="btn btn-primary btn-lg px-5 shadow-sm">
                    <i class="fa-solid fa-gears me-2"></i>Thực Hiện Trộn Dữ Liệu
                </button>
            </div>
        </form>
    </div>

    <div id="previewCard" class="card card-custom p-4 d-none">
        <h5 class="fw-bold text-secondary mb-3"><i class="fa-solid fa-table me-2"></i>Xem Trước Dữ Liệu Excel (DataTables)</h5>
        <div class="table-responsive">
            <table id="excelDataTable" class="table table-striped table-hover table-bordered w-100">
                <thead><tr id="tableHeader"></tr></thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.datatables.net/1.13.7/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.7/js/dataTables.bootstrap5.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<script>
    let dataTableInstance = null;

    document.getElementById('dataFile').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            if (jsonData.length > 0) {
                if (dataTableInstance) {
                    dataTableInstance.destroy();
                    \$('#excelDataTable').empty();
                }

                const headers = jsonData[0];
                const rows = jsonData.slice(1);

                let headerHtml = '<tr>';
                headers.forEach(h => headerHtml += `<th>${h || ''}</th>`);
                headerHtml += '</tr>';

                let bodyHtml = '';
                rows.forEach(r => {
                    if (r.length > 0) {
                        bodyHtml += '<tr>';
                        headers.forEach((_, idx) => bodyHtml += `<td>${r[idx] !== undefined ? r[idx] : ''}</td>`);
                        bodyHtml += '</tr>';
                    }
                });

                document.getElementById('excelDataTable').innerHTML = `<thead>${headerHtml}</thead><tbody>${bodyHtml}</tbody>`;
                document.getElementById('previewCard').classList.remove('d-none');

                dataTableInstance = \$('#excelDataTable').DataTable({
                    language: {
                        search: "<i class='fa-solid fa-magnifying-glass me-1'></i>Tìm kiếm:",
                        lengthMenu: "Hiển thị _MENU_ dòng",
                        info: "Đang xem _START_ đến _END_ trong tổng số _TOTAL_ dòng",
                        paginate: { first: "Đầu", last: "Cuối", next: "Sau", previous: "Trước" }
                    },
                    pageLength: 5,
                    lengthMenu: [5, 10, 25, 50]
                });
            }
        };
        reader.readAsArrayBuffer(file);
    });

    document.getElementById('mergeForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const templateFile = document.getElementById('templateFile').files[0];
        const dataFile = document.getElementById('dataFile').files[0];
        const filenameColumn = document.getElementById('filenameColumn').value.trim();
        const exportMode = document.getElementById('exportMode').value;
        const submitBtn = document.getElementById('submitBtn');

        const formData = new FormData();
        formData.append('template_file', templateFile);
        formData.append('data_file', dataFile);
        formData.append('export_mode', exportMode);
        if (filenameColumn) formData.append('filename_column', filenameColumn);

        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin me-2"></i>Đang xử lý dữ liệu...';

        try {
            const response = await fetch('/api/merge', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || 'Có lỗi xảy ra!');
            }

            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = exportMode === 'zip' ? 'Mail_Merge_Results.zip' : 'Mail_Merge_Merged.docx';
            document.body.appendChild(a);
            a.click();
            a.remove();

            alert('✅ Trộn biểu mẫu thành công!');
        } catch (err) {
            alert('❌ Lỗi: ' + err.message);
        } finally {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="fa-solid fa-gears me-2"></i>Thực Hiện Trộn Dữ Liệu';
        }
    });
</script>
</body>
</html>
"""

def replace_placeholders_in_doc(doc: Document, replacements: dict):
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
    for element in source_doc.element.body:
        target_doc.element.body.append(element)

@app.get("/", response_class=HTMLResponse)
def read_root():
    return HTMLResponse(content=HTML_CONTENT)

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
        data_bytes = await data_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)
        sheet = wb.active
        rows = list(sheet.iter_rows(values_only=True))

        if not rows or len(rows) < 2:
            raise HTTPException(status_code=400, detail="File Excel không có dữ liệu!")

        headers = [str(cell).strip() if cell is not None else f"Column_{i}" for i, cell in enumerate(rows[0])]
        data_rows = rows[1:]

        template_bytes = await template_file.read()

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