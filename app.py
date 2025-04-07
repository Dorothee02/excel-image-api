from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract_images():
    # Debug：印出 Content-Type
    print("Content-Type:", request.content_type)

    # Debug：列出收到的欄位名稱
    print("收到的 request.files keys：", list(request.files.keys()))

    # 檢查 Content-Type 是否正確
    if not request.content_type or not request.content_type.startswith("multipart/form-data"):
        return jsonify({"error": "Invalid content type"}), 400

    # 檢查是否有收到欄位名稱為 'file' 的檔案
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    # 儲存檔案到暫存資料夾
    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "file.xlsx")
    file.save(xlsx_path)

    # 解壓縮 xlsx
    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    # 讀取工作表名稱對照
    wb = load_workbook(xlsx_path)
    ws = wb.active
    name_map = {}
    for row in ws.iter_rows(min_row=6):
        row_index = row[0].row
        name = row[5].value
        if name:
            name_map[f'B{row_index}'] = str(name).strip()

    # 處理圖片
    media_path = os.path.join(unzip_path, "xl", "media")
    output_path = os.path.join(temp_dir, "output")
    os.makedirs(output_path, exist_ok=True)

    result = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path))
        for i, img_name in enumerate(images):
            row = i + 6
            cell = f'B{row}'
            new_name = f"{name_map.get(cell, 'unknown')}.jpeg"
            shutil.copyfile(
                os.path.join(media_path, img_name),
                os.path.join(output_path, new_name)
            )
            result.append({"filename": new_name})

    # 清理暫存資料夾
    shutil.rmtree(temp_dir)

    # 回傳結果
    return jsonify({"images": result})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
