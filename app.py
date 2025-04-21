from flask import Flask, request, jsonify, send_file
import os
import tempfile
import zipfile
from lxml import etree
import shutil
import glob

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
        
        # 創建輸出目錄
        output_dir = os.path.join(tmpdir, "output_images")
        os.makedirs(output_dir, exist_ok=True)
        
        try:
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
        except zipfile.BadZipFile:
            return jsonify({"error": "Invalid Excel file"}), 400
        
        # 檢查檔案是否存在
        drawing_folder = os.path.join(tmpdir, "xl/drawings")
        media_folder = os.path.join(tmpdir, "xl/media")
        
        drawing_exists = os.path.exists(drawing_folder)
        media_exists = os.path.exists(media_folder)
        
        if not drawing_exists or not media_exists:
            return jsonify({
                "status": "error",
                "message": "Excel file does not contain embedded images",
                "drawing_exists": drawing_exists,
                "media_exists": media_exists
            }), 400
        
        # 讀取關係文件，建立映射
        rels_map = {}
        rels_folder = os.path.join(tmpdir, "xl/drawings/_rels")
        if os.path.exists(rels_folder):
            for rels_file in os.listdir(rels_folder):
                if rels_file.endswith(".xml.rels"):
                    rels_path = os.path.join(rels_folder, rels_file)
                    try:
                        with open(rels_path, "rb") as f:
                            tree = etree.parse(f)
                            root = tree.getroot()
                            ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
                            for rel in root.findall(".//r:Relationship", namespaces=ns):
                                rid = rel.get("Id")
                                target = rel.get("Target")
                                if target and target.startswith("../media/"):
                                    image_file = target.split("/")[-1]
                                    rels_map[rid] = image_file
                    except Exception as e:
                        print(f"Error parsing relationships file {rels_file}: {e}")
        
        # 新增圖片插入位置 map
        image_cell_map = {}
        extracted_images = []
        
        if os.path.exists(drawing_folder):
            for fname in os.listdir(drawing_folder):
                if fname.endswith(".xml"):
                    drawing_xml_path = os.path.join(drawing_folder, fname)
                    try:
                        with open(drawing_xml_path, "rb") as f:
                            tree = etree.parse(f)
                            root = tree.getroot()
                            ns = {
                                "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                                "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
                            }
                            
                            # 處理 twoCellAnchor
                            for anchor in root.findall(".//xdr:twoCellAnchor", namespaces=ns):
                                try:
                                    from_col = int(anchor.find(".//xdr:from/xdr:col", namespaces=ns).text)
                                    from_row = int(anchor.find(".//xdr:from/xdr:row", namespaces=ns).text)
                                    cell_ref = colrow_to_cell(from_col, from_row)
                                    
                                    # 獲取圖片的引用ID
                                    blip = anchor.find(".//a:blip", namespaces=ns)
                                    if blip is not None:
                                        embed_rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                                        if embed_rid in rels_map:
                                            image_filename = rels_map[embed_rid]
                                            image_path = os.path.join(media_folder, image_filename)
                                            
                                            if os.path.exists(image_path):
                                                output_path = os.path.join(output_dir, f"{cell_ref}_{image_filename}")
                                                shutil.copy2(image_path, output_path)
                                                
                                                image_cell_map[image_filename] = cell_ref
                                                extracted_images.append({
                                                    "cell": cell_ref,
                                                    "original_filename": image_filename,
                                                    "saved_as": f"{cell_ref}_{image_filename}"
                                                })
                                except Exception as e:
                                    print(f"Error processing anchor: {e}")
                            
                            # 處理 oneCellAnchor
                            for anchor in root.findall(".//xdr:oneCellAnchor", namespaces=ns):
                                try:
                                    from_col = int(anchor.find(".//xdr:from/xdr:col", namespaces=ns).text)
                                    from_row = int(anchor.find(".//xdr:from/xdr:row", namespaces=ns).text)
                                    cell_ref = colrow_to_cell(from_col, from_row)
                                    
                                    # 獲取圖片的引用ID
                                    blip = anchor.find(".//a:blip", namespaces=ns)
                                    if blip is not None:
                                        embed_rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                                        if embed_rid in rels_map:
                                            image_filename = rels_map[embed_rid]
                                            image_path = os.path.join(media_folder, image_filename)
                                            
                                            if os.path.exists(image_path):
                                                output_path = os.path.join(output_dir, f"{cell_ref}_{image_filename}")
                                                shutil.copy2(image_path, output_path)
                                                
                                                image_cell_map[image_filename] = cell_ref
                                                extracted_images.append({
                                                    "cell": cell_ref,
                                                    "original_filename": image_filename,
                                                    "saved_as": f"{cell_ref}_{image_filename}"
                                                })
                                except Exception as e:
                                    print(f"Error processing anchor: {e}")
                    except Exception as e:
                        print(f"Failed to parse {fname}: {e}")
        
        # 創建ZIP壓縮包保存提取的圖片
        zip_output = os.path.join(tmpdir, "extracted_images.zip")
        with zipfile.ZipFile(zip_output, 'w') as zipf:
            for image_info in extracted_images:
                source_path = os.path.join(output_dir, image_info["saved_as"])
                zipf.write(source_path, arcname=image_info["saved_as"])
        
        # 返回結果
        return jsonify({
            "status": "success",
            "extracted_images": extracted_images,
            "cell_image_map": image_cell_map
        })

# 執行於 Zeabur 時固定用 port 8080
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
