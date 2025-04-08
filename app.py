from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64
from openpyxl import load_workbook
import re

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract_images():
    # 檢查是否有收到名為 'file' 的檔案
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    # 建立暫存資料夾並儲存檔案
    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "file.xlsx")
    file.save(xlsx_path)

    # 解壓縮 Excel 檔案
    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    # 讀取第 6 列開始的 F 欄（index 5）作為商品編號
    wb = load_workbook(xlsx_path)
    ws = wb.active
    product_ids = []
    for row in ws.iter_rows(min_row=6):
        product_ids.append(str(row[5].value).strip() if row[5].value else "unknown")

    # 取得圖片檔案路徑
    media_path = os.path.join(unzip_path, "xl", "media")
    output_images = []

    # 解析圖片檔名中的數字順序
    def extract_number(filename):
        match = re.search(r'(\d+)', filename)
        return int(match.group(1)) if match else 0

    if os.path.exists(media_path):
        # 處理所有 jpeg/jpg/png 格式的圖
        images = [img for img in os.listdir(media_path) if img.lower().endswith((".jpeg", ".jpg", ".png"))]
        images = sorted(images, key=extract_number)

        # 對應圖片與商品編號並回傳 base64 資料
        for i, img_name in enumerate(images):
            img_path = os.path.join(media_path, img_name)
            with open(img_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
            filename = f"{product_ids[i]}.jpeg" if i < len(product_ids) else f"unknown_{i}.jpeg"
            output_images.append({
                "filename": filename,
                "content": encoded,
                "mime_type": "image/jpeg"
            })

    # 清理暫存資料夾
    shutil.rmtree(temp_dir)

    return jsonify({"images": output_images})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
