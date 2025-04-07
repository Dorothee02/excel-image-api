from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract_images():
    print("Content-Type:", request.content_type)
    print("收到的 request.files keys：", list(request.files.keys()))

    if not request.content_type or not request.content_type.startswith("multipart/form-data"):
        return jsonify({"error": "Invalid content type"}), 400

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "file.xlsx")
    file.save(xlsx_path)

    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    wb = load_workbook(xlsx_path)
    ws = wb.active

    # 建立 row => name 的對應
    name_map = {}
    for row in ws.iter_rows(min_row=6):
        row_index = row[0].row
        name = row[5].value  # F欄是 index 5
        if name:
            name_map[row_index] = str(name).strip()

    # Debug：印出圖片實際貼在哪些格子
    for img in ws._images:
        try:
            anchor = img.anchor._from
            row = anchor.row + 1
            col = anchor.col + 1
            print(f"📸 圖片貼在第 {row} 列，第 {col} 欄")
        except Exception as e:
            print(f"⚠️ 無法取得圖片位置：{e}")

    # 處理圖片檔案
    media_path = os.path.join(unzip_path, "xl", "media")
    output_path = os.path.join(temp_dir, "output")
    os.makedirs(output_path, exist_ok=True)

    result = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path))
        for i, img_name in enumerate(images):
            row = i + 6  # 圖片順序假設從第6列起
            filename = f"{name_map.get(row, 'unknown')}.jpeg"
            shutil.copyfile(
                os.path.join(media_path, img_name),
                os.path.join(output_path, filename)
            )
            result.append({"filename": filename})

    shutil.rmtree(temp_dir)
    return jsonify({"images": result})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
