from flask import Flask, request, jsonify
import os
import tempfile
import zipfile

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    # 儲存在暫存目錄
    with tempfile.TemporaryDirectory() as tmpdir:
        filepath = os.path.join(tmpdir, file.filename)
        file.save(filepath)

        try:
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
        except zipfile.BadZipFile:
            return jsonify({"error": "Invalid Excel file"}), 400

        # 檢查需要的檔案
        drawing_xml = os.path.exists(os.path.join(tmpdir, "xl/drawings/drawing1.xml"))
        media_folder = os.path.exists(os.path.join(tmpdir, "xl/media"))
        worksheet_found = any(fname.endswith(".xml") for fname in os.listdir(os.path.join(tmpdir, "xl/worksheets")))

        return jsonify({
            "status": "success",
            "drawing1.xml exists": drawing_xml,
            "media folder exists": media_folder,
            "worksheet found": worksheet_found
        })

# 執行於 Zeabur 時固定用 port 8080
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
