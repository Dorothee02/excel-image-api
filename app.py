from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract_images():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "file.xlsx")
    file.save(xlsx_path)

    # 解壓縮 xlsx 檔
    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    # 讀取 Excel
    wb = load_workbook(xlsx_path)
    ws = wb.active

    # 建立 row -> F欄資料 的對照
    name_map = {}
    for row in ws.iter_rows(min_row=6):
        row_index = row[0].row
        name = row[5].value  # F欄是 index 5
        if name:
            name_map[row_index] = str(name).strip()

    # 取得圖片插入的列位置
    image_row_map = {}
    for i, img in enumerate(ws._images):
        try:
            anchor = img.anchor._from
            row = anchor.row + 1  # openpyxl row 從 0 起算
            image_row_map[i] = row
            print(f"📸 圖片 {i} 貼在第 {row} 列")
        except Exception as e:
            print(f"⚠️ 圖片 {i} 無法取得位置：{e}")

    # 處理圖片檔案
    media_path = os.path.join(unzip_path, "xl", "media")
    result = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path))
        for i, img_name in enumerate(images):
            row = image_row_map.get(i, i + 6)  # 如果無法定位，就假設是第 6+i 列
            filename = f"{name_map.get(row, 'unknown')}.jpeg"
            image_path = os.path.join(media_path, img_name)

            with open(image_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")

            result.append({
                "filename": filename,
                "content": encoded,
                "mime_type": "image/jpeg"
            })

    shutil.rmtree(temp_dir)
    return jsonify({"images": result})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
