from flask import Flask, request, jsonify
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import openpyxl
import os

app = Flask(__name__)

# XML namespaces
NAMESPACE_DEFS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main"
}

# 1. 取得特定工作表的 drawing XML 檔路徑列表
def get_sheet_drawing_paths(zf, sheet_idx=1):
    rels_path = f"xl/worksheets/_rels/sheet{sheet_idx}.xml.rels"
    if rels_path not in zf.namelist():
        return []
    tree = ET.fromstring(zf.read(rels_path))
    drawing_files = []
    for rel in tree.findall("Relationship", {"": ""}):
        if rel.attrib.get("Type", "").endswith("/drawing"):
            target = rel.attrib["Target"]  # e.g., '../drawings/drawing1.xml'
            path = os.path.normpath("xl/" + target.replace("../", ""))
            drawing_files.append(path)
    return drawing_files

# 2. 解析指定工作表的圖片 anchor（row, col, rId）
def parse_sheet_anchors(xlsx_bytes, sheet_idx=1):
    zf = zipfile.ZipFile(BytesIO(xlsx_bytes))
    drawing_paths = get_sheet_drawing_paths(zf, sheet_idx)
    anchors = []
    for dr in drawing_paths:
        xml = zf.read(dr)
        tree = ET.fromstring(xml)
        for tag in ("oneCellAnchor", "twoCellAnchor"):
            for anc in tree.findall(f"xdr:{tag}", NAMESPACE_DEFS):
                frm = anc.find("xdr:from", NAMESPACE_DEFS)
                row = int(frm.find("xdr:row", NAMESPACE_DEFS).text) + 1
                col = int(frm.find("xdr:col", NAMESPACE_DEFS).text) + 1
                blip = anc.find(".//a:blip", NAMESPACE_DEFS)
                if blip is not None:
                    rId = blip.attrib[f"{{{NAMESPACE_DEFS['r']}}}embed"]
                    anchors.append((row, col, rId))
    return anchors, zf

# 3. 建立 rId -> media 路徑映射
def build_media_map(zf):
    media = {}
    for name in zf.namelist():
        if name.startswith("xl/drawings/_rels/") and name.endswith(".rels"):
            tree = ET.fromstring(zf.read(name))
            drawing_path = name.replace("/_rels/", "/").replace(".rels", "")
            base_dir = os.path.dirname(drawing_path)
            for rel in tree.findall("Relationship"):
                rId = rel.attrib["Id"]
                target = rel.attrib["Target"]  # e.g., '../media/image1.png'
                media_path = os.path.normpath(os.path.join(base_dir, target))
                media[rId] = media_path
    return media

# 4. 讀取試算表，建立 row -> JAN code 映射
def load_jan_map(xlsx_bytes, jan_keyword="JAN"):
    wb = openpyxl.load_workbook(BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    jan_col = None
    header_row = None
    for r in range(1, 11):
        values = [cell.value for cell in ws[r]]
        for idx, val in enumerate(values, start=1):
            if val and jan_keyword.lower() in str(val).strip().lower():
                jan_col = idx
                header_row = r
                break
        if jan_col:
            break
    if not jan_col:
        raise ValueError("找不到包含 JAN 的欄位")
    jan_map = {}
    for row in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(row=row, column=jan_col).value
        if code:
            jan_map[row] = str(code)
    return jan_map

# HTTP 接口：接收 XLSX，回傳解析結果
@app.route("/extract-images", methods=["POST"])
def extract_images():
    f = request.files.get("file")
    if not f or not f.filename.endswith(".xlsx"):
        return "請上傳 .xlsx 檔案", 400
    data = f.read()
    anchors, zf = parse_sheet_anchors(data)
    jan_map = load_jan_map(data)
    media_map = build_media_map(zf)
    result = {}
    for row, col, rId in anchors:
        jan = jan_map.get(row) or f"unknown_{row}"
        img_path = media_map.get(rId)
        if img_path and img_path in zf.namelist():
            result[jan] = zf.read(img_path)
    return jsonify({"extracted": list(result.keys())})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
