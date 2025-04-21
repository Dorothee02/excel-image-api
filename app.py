from flask import Flask, request, jsonify, send_file
import os
import tempfile
import zipfile
from lxml import etree
import shutil
import openpyxl
from PIL import Image
import io

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    with tempfile.TemporaryDirectory() as tmpdir:
        filepath = os.path.join(tmpdir, file.filename)
        file.save(filepath)
        
        output_dir = os.path.join(tmpdir, "output_images")
        os.makedirs(output_dir, exist_ok=True)
        
        # 同時使用多種方法提取圖片
        extracted_images = []
        file_structure = []
        
        # 方法1: 使用 zipfile 直接提取媒體文件
        try:
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                # 獲取文件結構以便調試
                file_structure = zip_ref.namelist()
                zip_ref.extractall(tmpdir)
                
                # 查找媒體文件
                media_folder = os.path.join(tmpdir, "xl/media")
                if os.path.exists(media_folder):
                    for idx, media_file in enumerate(os.listdir(media_folder)):
                        media_path = os.path.join(media_folder, media_file)
                        output_path = os.path.join(output_dir, f"img_{idx+1}_{media_file}")
                        shutil.copy2(media_path, output_path)
                        
                        extracted_images.append({
                            "method": "direct_extraction",
                            "id": idx + 1,
                            "original_name": media_file,
                            "saved_as": f"img_{idx+1}_{media_file}"
                        })
        except Exception as e:
            print(f"Direct extraction error: {str(e)}")
        
        # 方法2: 使用 openpyxl 嘗試獲取圖片位置
        try:
            wb = openpyxl.load_workbook(filepath)
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                if hasattr(sheet, '_images'):
                    for idx, img in enumerate(sheet._images):
                        try:
                            img_data = img._data()
                            cell_ref = "unknown"
                            
                            if hasattr(img, 'anchor'):
                                anchor = img.anchor
                                if hasattr(anchor, '_from'):
                                    col = anchor._from.col
                                    row = anchor._from.row
                                    cell_ref = f"{chr(65 + col)}{row + 1}"
                            
                            # 保存圖片
                            img_filename = f"{sheet_name}_{cell_ref}_{idx}.png"
                            img_path = os.path.join(output_dir, img_filename)
                            
                            with open(img_path, 'wb') as f:
                                f.write(img_data)
                            
                            # 查看是否已經有相同圖片被提取
                            duplicate = False
                            for existing in extracted_images:
                                if existing.get("cell_ref") == cell_ref and existing.get("sheet") == sheet_name:
                                    duplicate = True
                                    break
                            
                            if not duplicate:
                                extracted_images.append({
                                    "method": "openpyxl",
                                    "sheet": sheet_name,
                                    "cell_ref": cell_ref,
                                    "saved_as": img_filename
                                })
                        except Exception as e:
                            print(f"Error extracting image with openpyxl: {str(e)}")
        except Exception as e:
            print(f"Openpyxl error: {str(e)}")
        
        # 方法3: 分析工作表中的關係文件
        try:
            rels_folder = os.path.join(tmpdir, "xl/worksheets/_rels")
            if os.path.exists(rels_folder):
                for rels_file in os.listdir(rels_folder):
                    if rels_file.endswith(".xml.rels"):
                        sheet_name = rels_file.split(".")[0]
                        rels_path = os.path.join(rels_folder, rels_file)
                        
                        with open(rels_path, "rb") as f:
                            tree = etree.parse(f)
                            root = tree.getroot()
                            ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
                            
                            for rel in root.findall(".//r:Relationship", namespaces=ns):
                                target = rel.get("Target")
                                if target and "../drawings/" in target:
                                    drawing_file = target.split("/")[-1]
                                    print(f"Found drawing reference: {sheet_name} -> {drawing_file}")
        except Exception as e:
            print(f"Relationship analysis error: {str(e)}")
        
        # 創建結果 ZIP 檔案
        if extracted_images:
            zip_output = os.path.join(tmpdir, "extracted_images.zip")
            with zipfile.ZipFile(zip_output, 'w') as zipf:
                for img_info in extracted_images:
                    img_filename = img_info.get("saved_as")
                    img_path = os.path.join(output_dir, img_filename)
                    if os.path.exists(img_path):
                        zipf.write(img_path, arcname=img_filename)
            
            return jsonify({
                "status": "success",
                "message": "Images extracted successfully",
                "extracted_count": len(extracted_images),
                "images": extracted_images
            })
        else:
            return jsonify({
                "status": "error",
                "message": "No images found in the Excel file",
                "file_structure": file_structure[:100]
            }), 404

@app.route('/test', methods=['GET'])
def test():
    return jsonify({"status": "API is running"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
