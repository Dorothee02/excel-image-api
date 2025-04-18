from flask import Flask, request, jsonify
import os
import tempfile
import zipfile
from lxml import etree

app = Flask(__name__)

# 工具函式：col/row 轉 Excel 格式（例如 A1, E6）
def colrow_to_cell(col, row):
    col_letter = ''
    while col >= 0:
        col_letter = chr(col % 26 + ord('A')) + col_letter
        col = col // 26 - 1
    return f"{col_letter}{row + 1}"

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

        # 檢查檔案是否存在
        drawing_xml = os.path.exists(os.path.join(tmpdir, "xl/drawings/drawing1.xml"))
        media_folder = os.path.exists(os.path.join(tmpdir, "xl/media"))
        worksheet_folder = os.path.join(tmpdir, "xl/worksheets")
        worksheet_found = os.path.exists(worksheet_folder) and any(fname.endswith(".xml") for fname in os.listdir(worksheet_folder))

        # 新增圖片插入位置 map
        image_cell_map = {}
        drawings_path = os.path.join(tmpdir, "xl/drawings")
        if os.path.exists(drawings_path):
            for fname in os.listdir(drawings_path):
                if fname.endswith(".xml"):
                    drawing_xml_path = os.path.join(drawings_path, fname)
                    try:
                        with open(drawing_xml_path, "rb") as f:
                            tree = etree.parse(f)
                            root = tree.getroot()

                            ns = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
                            anchors = root.findall(".//xdr:twoCellAnchor", namespaces=ns)

                            for idx, anchor in enumerate(anchors):
                                try:
                                    from_col = int(anchor.find(".//xdr:from/xdr:col", namespaces=ns).text)
                                    from_row = int(anchor.find(".//xdr:from/xdr:row", namespaces=ns).text)
                                    cell_ref = colrow_to_cell(from_col, from_row)
                                    image_filename = f"image{idx + 1}.png"
                                    image_cell_map[image_filename] = cell_ref
                                except:
                                    continue
                    except Exception as e:
                        print(f"Failed to parse {fname}: {e}")

        return jsonify({
            "status": "success",
            "drawing1.xml exists": drawing_xml,
            "media folder exists": media_folder,
            "worksheet found": worksheet_found,
            "cell_image_map": image_cell_map
        })

# 執行於 Zeabur 時固定用 port 8080
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
