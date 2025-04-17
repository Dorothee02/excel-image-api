from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64, xml.etree.ElementTree as ET
from openpyxl import load_workbook

app = Flask(__name__)

EMU_PER_CELL = 9525 * 65
EMU_PER_ROW = 9525 * 20

def get_center_cell(from_col, from_col_off, to_col, to_col_off,
                     from_row, from_row_off, to_row, to_row_off):
    center_x = (from_col + from_col_off / EMU_PER_CELL + to_col + to_col_off / EMU_PER_CELL) / 2
    center_y = (from_row + from_row_off / EMU_PER_ROW + to_row + to_row_off / EMU_PER_ROW) / 2
    return int(center_y), int(center_x)

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

    jan_col_index = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10)):
        for j, cell in enumerate(row):
            if cell.value and isinstance(cell.value, str) and ("JAN" in cell.value.upper()):
                jan_col_index = j
                break
        if jan_col_index is not None:
            break
    if jan_col_index is None:
        return jsonify({"error": "找不到 JAN 欄位"}), 400

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

    tree = ET.parse(drawing_xml)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
    output_images = []
    debug_log = []
    unknown_count = 1

    for anchor in tree.findall(".//a:twoCellAnchor", ns):
        from_tag = anchor.find("a:from", ns)
        to_tag = anchor.find("a:to", ns)
        if from_tag is None or to_tag is None:
            continue

        from_col = int(from_tag.find("a:col", ns).text)
        from_col_off = int(from_tag.find("a:colOff", ns).text)
        from_row = int(from_tag.find("a:row", ns).text)
        from_row_off = int(from_tag.find("a:rowOff", ns).text)

        to_col = int(to_tag.find("a:col", ns).text)
        to_col_off = int(to_tag.find("a:colOff", ns).text)
        to_row = int(to_tag.find("a:row", ns).text)
        to_row_off = int(to_tag.find("a:rowOff", ns).text)

        center_row, center_col = get_center_cell(
            from_col, from_col_off, to_col, to_col_off,
            from_row, from_row_off, to_row, to_row_off
        )

        pic = anchor.find("a:pic", ns)
        if pic is None:
            continue
        blip = pic.find(".//a:blip", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
        if blip is None:
            continue
        embed = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        img_file = rels_map.get(embed)
        full_img_path = os.path.join(media_path, img_file)
        if not os.path.exists(full_img_path):
            continue

        jan_value = f"unknown_{unknown_count}"
        if center_row is not None:
            row_idx = center_row + 1
            jan_cell = ws.cell(row=row_idx, column=jan_col_index + 1)
            if jan_cell.value:
                jan_value = str(jan_cell.value).strip()
            else:
                jan_value = f"row{row_idx}"
        else:
            unknown_count += 1

        with open(full_img_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")

        output_images.append({
            "filename": f"{jan_value}.jpeg",
            "content": encoded,
            "mime_type": "image/jpeg"
        })

        debug_log.append({
            "image": img_file,
            "center_row_col": (center_row, center_col),
            "assigned_name": jan_value
        })

    shutil.rmtree(temp_dir)
    return jsonify({
        "images": output_images,
        "debug": debug_log
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
