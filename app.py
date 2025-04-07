from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64
from openpyxl import load_workbook
import xml.etree.ElementTree as ET

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

    # 解壓縮 Excel 檔案
    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    # 讀取對應欄位的資料（例如 F 欄 = index 5）
    wb = load_workbook(xlsx_path)
    ws = wb.active
    name_map = {}
    for row in ws.iter_rows(min_row=6):
        row_index = row[0].row
        value = row[5].value
        if value:
            name_map[row_index] = str(value).strip()

    # 讀取圖片位置對應的 row（從 drawing1.xml）
    drawing_path = os.path.join(unzip_path, "xl", "drawings", "drawing1.xml")
    rels_path = os.path.join(unzip_path, "xl", "drawings", "_rels", "drawing1.xml.rels")
    media_path = os.path.join(unzip_path, "xl", "media")

    # 解析 rels 取得圖片對應的檔名
    rels_map = {}
    if os.path.exists(rels_path):
        rels_tree = ET.parse(rels_path)
        for rel in rels_tree.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rId = rel.attrib['Id']
            target = os.path.basename(rel.attrib['Target'])
            rels_map[rId] = target

    # 解析圖片位置與對應圖片
    img_order = []
    if os.path.exists(drawing_path):
        tree = ET.parse(drawing_path)
        ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
        for one in tree.findall(".//xdr:twoCellAnchor", ns):
            from_row_elem = one.find("xdr:from/xdr:row", ns)
            blip = one.find(".//xdr:blip", ns)
            if from_row_elem is not None and blip is not None:
                from_row = int(from_row_elem.text) + 1
                r_embed = blip.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
                img_name = rels_map.get(r_embed)
                if img_name:
                    print(f"drawing cell row {from_row} → {img_name}")
                    img_order.append((from_row, img_name))

    # 建立 output
    output_images = []
    for row_idx, img_name in img_order:
        product_id = name_map.get(row_idx, "unknown")
        img_path = os.path.join(media_path, img_name)
        if os.path.exists(img_path):
            with open(img_path, "rb") as img_file:
                encoded_string = base64.b64encode(img_file.read()).decode("utf-8")
            output_images.append({
                "filename": f"{product_id}.jpeg",
                "content": encoded_string,
                "mime_type": "image/jpeg"
            })
        else:
            print(f"⚠️ 找不到圖片: {img_path}")

    shutil.rmtree(temp_dir)
    return jsonify({"images": output_images})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
