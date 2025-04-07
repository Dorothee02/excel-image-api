from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract_images():
    print("Content-Type:", request.content_type)
    print("Received files:", request.files)

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "file.xlsx")
    file.save(xlsx_path)

    # 解壓縮 Excel
    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    # 讀取 F 欄（index 5）資料，從第 6 列開始
    wb = load_workbook(xlsx_path)
    ws = wb.active
    product_ids = []
    for row in ws.iter_rows(min_row=6):
        value = row[5].value  # F欄
        if value:
            product_ids.append(str(value).strip())

    # 取得圖片清單並排序
    media_path = os.path.join(unzip_path, "xl", "media")
    images = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path))

    # 依序配對圖片與商品編號
    output_images = []
    for i, img_name in enumerate(images):
        if i >= len(product_ids):
            break  # 避免圖片比資料多

        product_id = product_ids[i]
        img_path = os.path.join(media_path, img_name)
        if os.path.exists(img_path):
            with open(img_path, "rb") as img_file:
                encoded_string = base64.b64encode(img_file.read()).decode("utf-8")
            output_images.append({
                "filename": f"{product_id}.jpeg",
                "content": encoded_string,
                "mime_type": "image/jpeg"
            })

    shutil.rmtree(temp_dir)
    return jsonify({"images": output_images})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
