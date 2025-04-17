from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64, xml.etree.ElementTree as ET
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

    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    # 🔍 尋找欄位名稱含 JAN / ＪＡＮ 的欄位
    header_row = None
    jan_col_index = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10)):
        for j, cell in enumerate(row):
            if cell.value and isinstance(cell.value, str) and ("JAN" in cell.value.upper()):
                header_row = i + 1
                jan_col_index = j
                break
        if jan_col_index is not None:
            break
    if jan_col_index is None:
        return jsonify({"error": "找不到 JAN 或 ＪＡＮ 欄位"}), 400

    # 🔧 解析圖片與插入位置
    drawing_xml = os.path.join(unzip_path, "xl", "drawings", "drawing1.xml")
    rels_xml = os.path.join(unzip_path, "xl", "drawings", "_rels", "drawing1.xml.rels")
    media_path = os.path.join(unzip_path, "xl", "media")

    rels_map = {}
    if os.path.exists(rels_xml):
        rels_tree = ET.parse(rels_xml)
        for rel in rels_tree.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rId = rel.attrib['Id']
            target = rel.attrib['Target'].split('/')[-1]
            rels_map[rId] = target

    image_info = []
    if os.path.exists(drawing_xml):
        tree = ET.parse(drawing_xml)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
        for anchor in tree.findall(".//a:twoCellAnchor", ns):
            from_row = int(anchor.find("a:from/a:row", ns).text)
            to_row = int(anchor.find("a:to/a:row", ns).text)
            center_row = round((from_row + to_row) / 2)

            pic = anchor.find("a:pic", ns)
            if pic is not None:
                blip = pic.find(".//a:blip", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
                if blip is not None:
                    embed = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    filename = rels_map.get(embed)
                    image_info.append({"row": center_row, "filename": filename})

    image_info = sorted(image_info, key=lambda x: x["row"])
    output_images = []
    debug_log = []

    for info in image_info:
        row_idx = info["row"] + 1  # openpyxl 是 1-based
        img_file = info["filename"]
        full_img_path = os.path.join(media_path, img_file)

        if not os.path.exists(full_img_path):
            continue

        jan_cell = ws.cell(row=row_idx, column=jan_col_index + 1)
        cell_value = jan_cell.value
        if cell_value is None:
            jan_value = f"row{row_idx}"
        else:
            jan_value = str(cell_value).strip()

        with open(full_img_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")

        output_images.append({
            "filename": f"{jan_value}.jpeg",
            "content": encoded,
            "mime_type": "image/jpeg"
        })

        debug_log.append({
            "image": img_file,
            "row": row_idx,
            "jan_value": jan_value,
            "cell_raw": str(cell_value)
        })

    shutil.rmtree(temp_dir)

    return jsonify({
        "images": output_images,
        "debug": debug_log
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
