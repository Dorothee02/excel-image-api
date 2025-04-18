from flask import Flask, request, jsonify
import zipfile
import os
import tempfile
import xml.etree.ElementTree as ET

app = Flask(__name__)

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

        # 路徑定義
        drawing_path = os.path.join(tmpdir, "xl/drawings/drawing1.xml")
        rels_path = os.path.join(tmpdir, "xl/drawings/_rels/drawing1.xml.rels")
        media_path = os.path.join(tmpdir, "xl/media")
        worksheet_dir = os.path.join(tmpdir, "xl/worksheets")

        result = {
            "drawing1.xml exists": os.path.exists(drawing_path),
            "media folder exists": os.path.exists(media_path),
            "rels exists": os.path.exists(rels_path),
            "worksheet found": False,
            "cell_image_map": {}  # 最終目標對照表
        }

        # 嘗試抓出一個 worksheet 名稱（非必要）
        if os.path.exists(worksheet_dir):
            for fname in os.listdir(worksheet_dir):
                if fname.endswith(".xml"):
                    result["worksheet found"] = True
                    break

        # 準備讀 rels → rId 對應的圖片檔名
        rid_to_img = {}
        if os.path.exists(rels_path):
            tree = ET.parse(rels_path)
            root = tree.getroot()
            for rel in root:
                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target")  # 通常像 ../media/image4.png
                if rid and target and "media/" in target:
                    img_name = os.path.basename(target)
                    rid_to_img[rid] = img_name

        # 開始讀 drawing1.xml
        if os.path.exists(drawing_path):
            tree = ET.parse(drawing_path)
            root = tree.getroot()
            ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}

            cell_image_map = {}

            for anchor in root.findall('a:twoCellAnchor', ns):
                from_elem = anchor.find('a:from', ns)
                to_elem = anchor.find('a:to', ns)
                pic_elem = anchor.find('a:pic', ns)

                # 沒有圖片或位置資訊就略過
                if from_elem is None or to_elem is None or pic_elem is None:
                    continue

                # 算中心點 row、col
                row_start = int(from_elem.find('a:row', ns).text)
                row_end = int(to_elem.find('a:row', ns).text)
                row_center = round((row_start + row_end) / 2)

                col_start = int(from_elem.find('a:col', ns).text)
                col_end = int(to_elem.find('a:col', ns).text)
                col_center = round((col_start + col_end) / 2)

                col_letter = chr(ord('A') + col_center)
                cell_ref = f"{col_letter}{row_center + 1}"  # row 從 0 開始

                # 取得 r:embed → 轉成圖片名稱
                blip = pic_elem.find(".//a:blip", ns)
                if blip is None:
                    continue

                rId = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                image_filename = rid_to_img.get(rId, f"unknown_{rId}.png")

                cell_image_map[cell_ref] = image_filename

            result["cell_image_map"] = cell_image_map

        return jsonify(result)

# Zeabur port 設定
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
