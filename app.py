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

    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    # 讀取第 6 列開始的 F 欄（index 5）作為圖片檔名
    wb = load_workbook(xlsx_path)
    ws = wb.active
    product_ids = []
    for row in ws.iter_rows(min_row=6):
        product_ids.append(str(row[5].value).strip() if row[5].value else "unknown")

    # 處理 media 裡的圖片
    media_path = os.path.join(unzip_path, "xl", "media")
    output_images = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path), key=lambda x: int(''.join(filter(str.isdigit, x))))
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

    shutil.rmtree(temp_dir)
    return jsonify({"images": output_images})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
