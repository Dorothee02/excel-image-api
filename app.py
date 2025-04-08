from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64, re
from openpyxl import load_workbook

app = Flask(__name__)

def natural_sort_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]

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

    wb = load_workbook(xlsx_path)
    ws = wb.active

    media_path = os.path.join(unzip_path, "xl", "media")
    output_images = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path), key=natural_sort_key)
        image_count = len(images)

        # 讀取第 6 列開始、F 欄的值，抓到跟圖片一樣多的資料
        product_ids = []
        for i in range(image_count):
            cell_value = ws.cell(row=6 + i, column=6).value  # column=6 是 F 欄
            product_ids.append(str(cell_value).strip() if cell_value else "unknown")

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
