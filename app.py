from flask import Flask, request, jsonify
import zipfile
import os
import tempfile
import xml.etree.ElementTree as ET
import base64

app = Flask(__name__)

# 🧠 升級：支援 A ~ ZZ 甚至 AAA 欄位命名
def col_index_to_letter(index):
    letters = ''
    while index >= 0:
        letters = chr(index % 26 + ord('A')) + letters
        index = index // 26 - 1
    return letters

@app.route('/upload', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    with tempfile.TemporaryDirectory() as tmpdir:
        filepath = os.path.join(tmpdir, file.filename)
        file.save(filepath)

        try:
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
        except zipfile.BadZipFile:
            return jsonify({"error": "Invalid Excel file"}), 400

        drawing_path = os.path.join(tmpdir, "xl/drawings/drawing1.xml")
        media_path = os.path.join(tmpdir, "xl/media")
        worksheet_dir = os.path.join(tmpdir, "xl/worksheets")

        result = {
            "cell_image_map": {},
            "drawing1.xml exists": os.path.exists(drawing_path),
            "media folder exists": os.path.exists(media_path),
            "worksheet found": False
        }

        if os.path.exists(worksheet_dir):
            for fname in os.listdir(worksheet_dir):
                if fname.endswith(".xml"):
                    result["worksheet found"] = True
                    break

        # 🔍 收集圖片清單（按檔名排序）
        image_files = sorted(
            [f for f in os.listdir(media_path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        ) if os.path.exists(media_path) else []

        # 📍 抓出圖片插入位置並用儲存格名稱命名
        if os.path.exists(drawing_path):
            ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
            tree = ET.parse(drawing_path)
            root = tree.getroot()

            image_map = {}
            image_index = 0

            for anchor in root.findall('a:twoCellAnchor', ns):
                from_elem = anchor.find('a:from', ns)
                to_elem = anchor.find('a:to', ns)
                if from_elem is None or to_elem is None:
                    continue

                row_start = int(from_elem.find('a:row', ns).text)
                row_end = int(to_elem.find('a:row', ns).text)
                row_center = round((row_start + row_end) / 2)

                col_start = int(from_elem.find('a:col', ns).text)
                col_end = int(to_elem.find('a:col', ns).text)
                col_center = round((col_start + col_end) / 2)

                col_letter = col_index_to_letter(col_center)
                cell_ref = f"{col_letter}{row_center + 1}"

                # 🖼️ 對應圖片順序 → base64 + 附副檔名
                if image_index < len(image_files):
                    image_path = os.path.join(media_path, image_files[image_index])
                    with open(image_path, "rb") as img_file:
                        img_data = img_file.read()
                        ext = os.path.splitext(image_files[image_index])[-1].replace('.', '')
                        base64_img = f"data:image/{ext};base64," + base64.b64encode(img_data).decode('utf-8')
                        image_map[cell_ref] = base64_img
                    image_index += 1

            result["cell_image_map"] = image_map

        return jsonify(result)

# 🌐 Zeabur 要用 port 8080
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
