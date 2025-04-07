from flask import Flask, request, jsonify
import os
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET
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

    # 讀取 Excel
    wb = load_workbook(xlsx_path)
    ws = wb.active

    # 解析圖片儲存格對應資料
    drawing_path = os.path.join(unzip_path, "xl", "drawings", "drawing1.xml")
    if not os.path.exists(drawing_path):
        return jsonify({"error": "drawing1.xml not found"}), 500

    tree = ET.parse(drawing_path)
    root = tree.getroot()

    ns = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    }

    image_cells = []
    for anchor in root.findall("xdr:twoCellAnchor", ns):
        from_node = anchor.find("xdr:from", ns)
        row = int(from_node.find("xdr:row", ns).text)
        col = int(from_node.find("xdr:col", ns).text)
        image_cells.append((row, col))

    # 讀取第 row+1 列、第 F 欄 (index = 5) 作為圖片檔名依據
    name_map = {}
    for row, col in image_cells:
        excel_row = row + 1  # zero-based to 1-based row number
        cell_value = ws.cell(row=excel_row, column=6).value  # F欄 = 第6欄
        name_map[row] = str(cell_value).strip() if cell_value else "unknown"

    # 搬圖片
    media_path = os.path.join(unzip_path, "xl", "media")
    output_path = os.path.join(temp_dir, "output")
    os.makedirs(output_path, exist_ok=True)

    result = []
    if os.path.exists(media_path):
        images = sorted(os.listdir(media_path))
        for i, img_name in enumerate(images):
            # 如果圖片數量比儲存格還多，會抓不到對應 row
            if i < len(image_cells):
                row, _ = image_cells[i]
                new_name = f"{name_map.get(row, 'unknown')}.jpeg"
            else:
                new_name = f"unknown_{i}.jpeg"

            shutil.copyfile(
                os.path.join(media_path, img_name),
                os.path.join(output_path, new_name)
            )
            result.append({"filename": new_name})

    shutil.rmtree(temp_dir)
    return jsonify({"images": result})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
