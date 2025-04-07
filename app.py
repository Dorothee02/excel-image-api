from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil
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

    wb = load_workbook(xlsx_path)
    ws = wb.active
    name_map = {}
    for row in ws.iter_rows(min_row=6):
        row_index = row[0].row
        name = row[5].value
        if name:
            name_map[f'B{row_index}'] = str(name).strip()

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

    shutil.rmtree(temp_dir)
    return jsonify({"images": result})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
