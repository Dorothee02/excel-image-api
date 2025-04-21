from flask import Flask, request, jsonify
import os
import tempfile
import zipfile
from lxml import etree
import shutil
import re

app = Flask(__name__)

def find_cell_reference_in_worksheet(worksheet_path, image_names):
    """分析工作表 XML 尋找圖片引用"""
    try:
        with open(worksheet_path, 'rb') as f:
            content = f.read().decode('utf-8', errors='ignore')
            
        # 尋找例如 <drawing r:id="rId1"/> 的標籤
        drawing_refs = re.findall(r'<drawing r:id="(rId\d+)"/>', content)
        if not drawing_refs:
            return None
            
        # 尋找儲存格位置標記
        cell_positions = re.findall(r'<c r="([A-Z]+\d+)"', content)
        
        # 分析內容尋找圖片和儲存格的關聯
        # 這部分可能需要根據實際 XML 結構調整
        return cell_positions
    except Exception as e:
        print(f"Error analyzing worksheet: {str(e)}")
        return None

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
        
        try:
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
                file_list = zip_ref.namelist()
                
            # 步驟1: 找出所有工作表
            worksheet_files = []
            worksheet_folder = os.path.join(tmpdir, "xl/worksheets")
            if os.path.exists(worksheet_folder):
                worksheet_files = [f for f in os.listdir(worksheet_folder) if f.endswith('.xml')]
            
            # 步驟2: 提取媒體文件
            media_folder = os.path.join(tmpdir, "xl/media")
            extracted_images = []
            media_files = []
            
            if os.path.exists(media_folder):
                media_files = os.listdir(media_folder)
                for idx, media_file in enumerate(media_files):
                    media_path = os.path.join(media_folder, media_file)
                    output_path = os.path.join(output_dir, f"img_{idx+1}_{media_file}")
                    shutil.copy2(media_path, output_path)
                    
                    # 初始化圖片信息
                    image_info = {
                        "id": idx + 1,
                        "method": "direct_extraction",
                        "original_name": media_file,
                        "saved_as": f"img_{idx+1}_{media_file}"
                    }
                    extracted_images.append(image_info)
            
            # 步驟3: 高級定位 - 分析各工作表 XML
            for ws_file in worksheet_files:
                ws_path = os.path.join(worksheet_folder, ws_file)
                
                # 分析工作表 XML 尋找圖片引用
                cell_refs = find_cell_reference_in_worksheet(ws_path, media_files)
                if cell_refs:
                    print(f"Found potential cell references in {ws_file}: {cell_refs[:5]}...")
            
            # 步驟4: 嘗試通過關係文件定位
            rels_folder = os.path.join(tmpdir, "xl/worksheets/_rels")
            if os.path.exists(rels_folder):
                for rels_file in os.listdir(rels_folder):
                    if rels_file.endswith(".xml.rels"):
                        try:
                            rels_path = os.path.join(rels_folder, rels_file)
                            sheet_name = rels_file.split(".")[0]
                            
                            with open(rels_path, "rb") as f:
                                tree = etree.parse(f)
                                root = tree.getroot()
                                ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
                                
                                for rel in root.findall(".//r:Relationship", namespaces=ns):
                                    rel_id = rel.get("Id")
                                    target = rel.get("Target")
                                    rel_type = rel.get("Type")
                                    
                                    # 檢查是否為與繪圖相關的關係
                                    if "drawing" in target.lower():
                                        print(f"Drawing relationship found: {sheet_name} -> {rel_id} -> {target}")
                                        
                                        # 嘗試找到對應的工作表XML
                                        sheet_xml_path = os.path.join(worksheet_folder, f"{sheet_name}.xml")
                                        if os.path.exists(sheet_xml_path):
                                            with open(sheet_xml_path, "rb") as sheet_file:
                                                sheet_content = sheet_file.read().decode('utf-8', errors='ignore')
                                                # 尋找使用此關係ID的位置
                                                draw_tag = f'<drawing r:id="{rel_id}"/>'
                                                if draw_tag in sheet_content:
                                                    # 分析附近的標籤找出儲存格位置
                                                    cell_tags = re.findall(r'<c r="([A-Z]+\d+)"[^>]*>(?:.*?)</c>', 
                                                                          sheet_content, re.DOTALL)
                                                    if cell_tags:
                                                        print(f"Possible cell references: {cell_tags[:5]}")
                        except Exception as e:
                            print(f"Error analyzing relationship file {rels_file}: {e}")
            
            # 步驟5: 嘗試從content_types.xml獲取更多信息
            content_types_path = os.path.join(tmpdir, "[Content_Types].xml")
            if os.path.exists(content_types_path):
                try:
                    with open(content_types_path, "rb") as f:
                        tree = etree.parse(f)
                        root = tree.getroot()
                        
                        # 尋找與圖片相關的內容類型
                        for override in root.findall(".//{http://schemas.openxmlformats.org/package/2006/content-types}Override"):
                            part_name = override.get("PartName")
                            content_type = override.get("ContentType")
                            
                            if "drawing" in part_name.lower():
                                print(f"Drawing content found: {part_name} -> {content_type}")
                except Exception as e:
                    print(f"Error analyzing content types: {e}")
            
            # 創建 ZIP 檔案來傳送所有提取的圖片
            zip_output = os.path.join(tmpdir, "extracted_images.zip")
            with zipfile.ZipFile(zip_output, 'w') as zipf:
                for img_info in extracted_images:
                    img_path = os.path.join(output_dir, img_info["saved_as"])
                    zipf.write(img_path, arcname=img_info["saved_as"])
            
            # 返回結果
            return jsonify({
                "status": "success",
                "message": "Images extracted successfully",
                "extracted_count": len(extracted_images),
                "images": extracted_images,
                "file_structure": [f for f in file_list if f.startswith("xl/") and not f.endswith("/")][:20]
            })
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            return jsonify({
                "error": str(e),
                "details": error_details,
                "file_structure": file_list[:50] if 'file_list' in locals() else []
            }), 500

@app.route('/test', methods=['GET'])
def test():
    return jsonify({"status": "API is running"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
