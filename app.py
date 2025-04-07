from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64
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
    name_map = {}
    for row in ws.iter_rows(min_row=6):
        row_index = row[0].row
        name = row[5].value
        if name:
            name_map[f'B{row_index}'] = str(name).strip()

    media_path = os.path.join(unzip_path, "xl", "media")

    result = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path))
        for i, img_name in enumerate(images):
            row = i + 6
            cell = f'B{row}'
            new_name = f"{name_map.get(cell, 'unknown')}.jpeg"
            image_path = os.path.join(media_path, img_name)

            # 讀取圖片並轉成 base64
            with open(image_path, "rb") as img_file:
                encoded_string = base64.b64encode(img_file.read()).decode("utf-8")

            result.append({
                "filename": new_name,
                "content": encoded_string,
                "mime_type": "image/jpeg"
            })

    shutil.rmtree(temp_dir)
    return jsonify({"images": result})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
