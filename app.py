from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64, re
from openpyxl import load_workbook
import xml.etree.ElementTree as ET

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

    # 抓取 F 欄資料 (商品編號)
    wb = load_workbook(xlsx_path)
    ws = wb.active
    product_ids = []
    for row in ws.iter_rows(min_row=6):
        value = row[5].value  # F 欄 index = 5
        if value:
            product_ids.append(str(value).strip())

    # 解析圖片順序
    drawing_path = os.path.join(unzip_path, "xl", "drawings", "drawing1.xml")
    rels_path = os.path.join(unzip_path, "xl", "drawings", "_rels", "drawing1.xml.rels")
    media_path = os.path.join(unzip_path, "xl", "media")

    rels_map = {}
    if os.path.exists(rels_path):
        rels_tree = ET.parse(rels_path)
        for rel in rels_tree.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rId = rel.attrib['Id']
            target = os.path.basename(rel.attrib['Target'])
            rels_map[rId] = target

    # 抓圖形名稱並排序（照圖 1、圖 2、圖 8... 的數字遞增）
    image_seq = []
    if os.path.exists(drawing_path):
        tree = ET.parse(drawing_path)
        ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
        for anchor in tree.findall(".//xdr:twoCellAnchor", ns):
            docPr = anchor.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr", ns)
            blip = anchor.find(".//xdr:blip", ns)
            if docPr is not None and blip is not None:
                name = docPr.attrib.get("name")  # e.g., "图 8"
                match = re.search(r"\\d+", name)
                if match:
                    number = int(match.group())
                    embed = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    img_file = rels_map.get(embed)
                    if img_file:
                        image_seq.append((number, img_file))

    image_seq.sort(key=lambda x: x[0])  # 根據圖形編號排序

    # 照順序配對 F 欄的商品編號
    output_images = []
    for idx, (img_no, img_filename) in enumerate(image_seq):
        if idx >= len(product_ids):
            break
        img_path = os.path.join(media_path, img_filename)
        if os.path.exists(img_path):
            with open(img_path, "rb") as img_file:
                encoded_string = base64.b64encode(img_file.read()).decode("utf-8")
            output_images.append({
                "filename": f"{product_ids[idx]}.jpeg",
                "content": encoded_string,
                "mime_type": "image/jpeg"
            })

    shutil.rmtree(temp_dir)
    return jsonify({"images": output_images})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
