from flask import Flask, request, jsonify, send_file
import os
import tempfile
import zipfile
from lxml import etree
import shutil
import pandas as pd
import openpyxl
import re

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
        
        try:
            # 步驟1: 使用openpyxl讀取Excel檔案結構
            wb = openpyxl.load_workbook(filepath)
            
            # 步驟2: 找到JAN欄位的列號
            jan_column = None
            jan_header_row = None
            
            for sheet in wb.worksheets:
                for row_idx in range(1, min(20, sheet.max_row + 1)):  # 搜尋前20行
                    for col_idx in range(1, min(20, sheet.max_column + 1)):  # 搜尋前20列
                        cell_value = sheet.cell(row=row_idx, column=col_idx).value
                        if cell_value and isinstance(cell_value, str) and "JAN" in cell_value:
                            jan_column = col_idx
                            jan_header_row = row_idx
                            break
                    if jan_column:
                        break
                if jan_column:
                    break
            
            if not jan_column:
                return jsonify({"error": "Could not find JAN column in the Excel file"}), 400
            
            # 步驟3: 解壓縮Excel檔案
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
            
            # 步驟4: 提取圖片和位置資訊
            media_folder = os.path.join(tmpdir, "xl/media")
            if not os.path.exists(media_folder):
                return jsonify({"error": "No images found in the Excel file"}), 400
            
            # 步驟5: 嘗試使用openpyxl獲取圖片位置
            image_positions = []
            
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                if hasattr(sheet, '_images'):
                    for img in sheet._images:
                        try:
                            if hasattr(img, 'anchor') and hasattr(img.anchor, '_from'):
                                row = img.anchor._from.row + 1  # Excel行號從1開始
                                col = img.anchor._from.col
                                col_letter = openpyxl.utils.get_column_letter(col + 1)
                                
                                # 獲取對應行的JAN碼
                                jan_value = sheet.cell(row=row, column=jan_column).value
                                
                                image_positions.append({
                                    "row": row,
                                    "col": col_letter,
                                    "position": f"{col_letter}{row}",
                                    "jan_value": jan_value if jan_value else f"unknown_row{row}"
                                })
                        except Exception as e:
                            print(f"Error getting image position: {e}")
            
            # 步驟6: 如果openpyxl無法獲取位置，嘗試分析drawing關係
            if not image_positions:
                # 實現替代方案: 提取所有圖片，並根據檔案順序猜測行號
                media_files = os.listdir(media_folder)
                media_files.sort()  # 按名稱排序
                
                # 讀取Excel數據
                df = pd.read_excel(filepath)
                
                # 識別JAN列
                jan_col_name = None
                for col in df.columns:
                    if "JAN" in str(col):
                        jan_col_name = col
                        break
                
                if jan_col_name:
                    # 假設圖片順序與Excel行順序相同
                    for idx, media_file in enumerate(media_files):
                        # 除去header行
                        row_idx = min(idx + jan_header_row + 1, len(df) + jan_header_row)
                        
                        if row_idx - jan_header_row <= len(df):
                            jan_value = str(df.iloc[row_idx - jan_header_row - 1][jan_col_name])
                            if pd.isna(jan_value) or jan_value == "nan":
                                jan_value = f"unknown_row{row_idx}"
                        else:
                            jan_value = f"unknown_row{row_idx}"
                        
                        image_positions.append({
                            "row": row_idx,
                            "method": "estimated_by_order",
                            "jan_value": jan_value
                        })
            
            # 步驟7: 提取和重命名圖片
            extracted_images = []
            
            for idx, media_file in enumerate(os.listdir(media_folder)):
                media_path = os.path.join(media_folder, media_file)
                
                # 確定圖片對應的JAN值
                jan_value = "unknown"
                position = "unknown"
                
                if idx < len(image_positions):
                    jan_value = image_positions[idx]["jan_value"]
                    position = image_positions[idx].get("position", f"row{image_positions[idx]['row']}")
                
                # 建立新檔名
                new_filename = f"{jan_value}_{media_file}"
                output_path = os.path.join(output_dir, new_filename)
                
                # 複製圖片到輸出目錄
                shutil.copy2(media_path, output_path)
                
                extracted_images.append({
                    "original_name": media_file,
                    "position": position,
                    "row": image_positions[idx]["row"] if idx < len(image_positions) else "unknown",
                    "jan_value": jan_value,
                    "saved_as": new_filename
                })
            
            # 步驟8: 建立ZIP檔案
            zip_output = os.path.join(tmpdir, "extracted_images.zip")
            with zipfile.ZipFile(zip_output, 'w') as zipf:
                for img_info in extracted_images:
                    img_path = os.path.join(output_dir, img_info["saved_as"])
                    zipf.write(img_path, arcname=img_info["saved_as"])
            
            return jsonify({
                "status": "success",
                "message": "Images extracted and renamed with JAN codes",
                "jan_column_found": f"Column {openpyxl.utils.get_column_letter(jan_column)} (Header at row {jan_header_row})",
                "extracted_images": extracted_images
            })
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            return jsonify({
                "status": "error",
                "message": str(e),
                "details": error_details
            }), 500

@app.route('/test', methods=['GET'])
def test():
    return jsonify({"status": "API is running"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
