from flask import Flask, request, jsonify
import zipfile
import os
import tempfile

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_excel():
    # ✅ 檢查是否有收到檔案
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    # ✅ 建立一個臨時資料夾來儲存與處理上傳的 Excel 檔案
    with tempfile.TemporaryDirectory() as tmpdir:
        filepath = os.path.join(tmpdir, file.filename)
        file.save(filepath)

        # ✅ 解壓縮 .xlsx 檔案（其實是 zip 格式）
        try:
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
        except zipfile.BadZipFile:
            return jsonify({"error": "Invalid Excel file"}), 400

        # ✅ 檢查是否存在圖片位置設定檔 drawing1.xml
        drawing_path = os.path.join(tmpdir, "xl/drawings/drawing1.xml")
        drawing_exists = os.path.exists(drawing_path)

        # ✅ 檢查是否存在圖片實體檔案（media資料夾）
        media_path = os.path.join(tmpdir, "xl/media")
        media_exists = os.path.exists(media_path)

        # ✅ 嘗試找到第一個工作表的 XML 檔（名稱不固定）
        worksheet_path = None
        worksheet_dir = os.path.join(tmpdir, "xl/worksheets")
        if os.path.exists(worksheet_dir):
            for fname in os.listdir(worksheet_dir):
                if fname.endswith(".xml"):
                    worksheet_path = os.path.join(worksheet_dir, fname)
                    break

        # ✅ 回傳檔案檢查結果
        return jsonify({
            "drawing1.xml exists": drawing_exists,
            "media folder exists": media_exists,
            "worksheet found": worksheet_path is not None
        })

# ✅ Zeabur 要使用 port 8080，記得要綁定 0.0.0.0 才能對外開放
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
