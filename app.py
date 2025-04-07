from flask import Flask, request, jsonify
import os, zipfile, tempfile, shutil, base64
from openpyxl import load_workbook
import xml.etree.ElementTree as ET

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract_images():
    print("Content-Type:", request.content_type)
    print("Received files:", request.files)

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    temp_dir = tempfile.mkdtemp()
    xlsx_path = os.path.join(temp_dir, "file.xlsx")
    file.save(xlsx_path)

    # 解壓縮 Excel
    unzip_path = os.path.join(temp_dir, "unzipped")
    os.makedirs(unzip_path, exist_ok=True)
    with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_path)

    # 檢查路徑狀況
    print("檢查 unzip_path:", os.path.exists(unzip_path))
    print("列出 unzip_path 下的檔案:", os.listdir(unzip_path))

    # 看看 xl/drawings 是否存在
    drawings_dir = os.path.join(unzip_path, "xl", "drawings")
    print("drawings_dir 存在？", os.path.exists(drawings_dir))
    if os.path.exists(drawings_dir):
        print("drawings_dir 檔案列表:", os.listdir(drawings_dir))

    # 看看 media 裡有沒有圖片
    media_path = os.path.join(unzip_path, "xl", "media")
    print("media_path 存在？", os.path.exists(media_path))
    if os.path.exists(media_path):
        print("media_path 圖片檔案列表:", os.listdir(media_path))

    shutil.rmtree(temp_dir)
    return jsonify({"images": []})
