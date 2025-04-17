from flask import Flask, request, jsonify
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import openpyxl
from openpyxl import load_workbook
import os

app = Flask(__name__)

# XML 命名空間定義
NAMESPACE_DEFS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main"
}

# 1. 取得指定工作表的 drawing 路徑
def get_sheet_drawing_paths(zf, sheet_idx=1):
    rels_path = f"xl/worksheets/_rels/sheet{sheet_idx}.xml.rels"
    if rels_path not in zf.namelist():
        return []
    tree = ET.fromstring(zf.read(rels_path))
    drawing_files = []
    for rel in tree.findall("Relationship", {"": ""}):
        if rel.attrib.get("Type", "").endswith("/drawing"):
            target = rel.attrib["Target"]
            path = os.path.normpath("xl/" + target.replace("../", ""))
            drawing_files.append(path)
    return drawing_files

# 2. 解析指定工作表的 anchor (row, col, rId)
def parse_sheet_anchors(xlsx_bytes, sheet_idx=1):
    zf = zipfile.ZipFile(BytesIO(xlsx_bytes))
    drawing_paths = get_sheet_drawing_paths(zf, sheet_idx)
    anchors = []
    for dr in drawing_paths:
        tree = ET.fromstring(zf.read(dr))
        for tag in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor"):
            for anc in tree.findall(f"xdr:{tag}", NAMESPACE_DEFS):
                frm = anc.find("xdr:from", NAMESPACE_DEFS)
                if not frm:
                    continue
                row = int(frm.find("xdr:row", NAMESPACE_DEFS).text) + 1
                col = int(frm.find("xdr:col", NAMESPACE_DEFS).text) + 1
                blip = anc.find(".//a:blip", NAMESPACE_DEFS)
                if blip is not None:
                    rId = blip.attrib[f"{{{NAMESPACE_DEFS['r']}}}embed"]
                    anchors.append((row, col, rId))
    return anchors, zf

# 3. 建立 rId -> media 路徑對應
def build_media_map(zf):
    media = {}
    for name in zf.namelist():
        if name.startswith("xl/drawings/_rels/") and name.endswith(".rels"):
            tree = ET.fromstring(zf.read(name))
            drawing_path = name.replace("/_rels/", "/").replace(".rels", "")
            base_dir = os.path.dirname(drawing_path)
            for rel in tree.findall("Relationship"):
                rId = rel.attrib["Id"]
                target = rel.attrib["Target"]
                media_path = os.path.normpath(os.path.join(base_dir, target.replace('../','')))
                media[rId] = media_path
    return media

# 4. 建立 row -> JAN code 映射
def load_jan_map(xlsx_bytes, jan_keyword="JAN"):
    wb = openpyxl.load_workbook(BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    jan_col = None
    header_row = None
    for r in range(1, 11):
        vals = [cell.value for cell in ws[r]]
        for idx, val in enumerate(vals, start=1):
            if val and jan_keyword.lower() in str(val).strip().lower():
                jan_col, header_row = idx, r
                break
        if jan_col:
            break
    if not jan_col:
        raise ValueError("前10列找不到包含 JAN 的欄位")
    jan_map = {}
    for row in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(row=row, column=jan_col).value
        if code:
            jan_map[row] = str(code)
    return jan_map

# 備援：使用 openpyxl 內建 ws._images
def extract_with_openpyxl(xlsx_bytes, sheet_idx=1):
    wb = load_workbook(BytesIO(xlsx_bytes), data_only=True)
    ws = wb.worksheets[sheet_idx-1]
    jan_map = load_jan_map(xlsx_bytes)
    result = []
    for img in ws._images:
        row = img.anchor._from.row + 1
        jan = jan_map.get(row) or f"unknown_{row}"
        path = img.ref  # e.g. 'xl/media/image1.png'
        ext = os.path.splitext(path)[1]
        result.append(f"{jan}{ext}")
    return result

# 主處理函式
@app.route("/extract-images", methods=["POST"])
def extract_images():
    sheet_idx = int(request.args.get('sheet', 1))
    f = request.files.get("file")
    if not f or not f.filename.endswith(".xlsx"):
        return "請上傳 .xlsx 檔案", 400
    data = f.read()
    app.logger.info(f"Received: {f.filename}, sheet={sheet_idx}, size={len(data)} bytes")

    # 嘗試 ZIP+XML 解析
    anchors, zf = parse_sheet_anchors(data, sheet_idx)
    app.logger.info(f"Parsed anchors: {anchors}")
    jan_map = load_jan_map(data)
    media_map = build_media_map(zf)
    extracted = []
    for row, col, rId in anchors:
        jan = jan_map.get(row) or f"unknown_{row}"
        img_path = media_map.get(rId)
        if img_path and img_path in zf.namelist():
            ext = os.path.splitext(img_path)[1]
            extracted.append(f"{jan}{ext}")
    # 若無結果，fallback to openpyxl
    if not extracted:
        app.logger.info("No anchors found, falling back to openpyxl._images")
        extracted = extract_with_openpyxl(data, sheet_idx)
        app.logger.info(f"Fallback extracted: {extracted}")

    return jsonify({"extracted": extracted})

# 別名路由
@app.route("/extract", methods=["POST"])
def extract_images_alias():
    return extract_images()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
