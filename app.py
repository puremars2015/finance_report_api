from flask import Flask, request, jsonify, render_template, send_from_directory
import os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # 启用CORS

sourcefile = 'source.xlsx'
convertedfile = 'converted.xlsx'

# 设置文件上传的目录
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/convert')
def index():
    return render_template('upload_service.html',  error=None)

@app.route('/convert-with-auto-file-name')
def convert():
    return render_template('convert_service.html',  error=None)

@app.route('/')
def convert_excel_full_step():
    return render_template('convert_excel_service.html',  error=None)

@app.route('/upload', methods=['POST'])
def upload_file():
    # 检查请求中是否包含文件
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    
    # 如果用户没有选择文件
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    if file:
        # 保存文件到指定目录
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        return jsonify({"message": "File uploaded successfully", "filename": file.filename}), 200
    
@app.route('/upload-source', methods=['POST'])
def upload_source_file():
    # 检查请求中是否包含文件
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    
    # 如果用户没有选择文件
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    if file:
        # 保存文件到指定目录
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], sourcefile)
        file.save(filepath)
        return jsonify({"message": "File uploaded successfully", "filename": sourcefile}), 200

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found"}), 404
    
@app.route('/download-converted', methods=['GET'])
def download_converted_file():
    try:
        from converter import convert
        convert(os.path.join(app.config['UPLOAD_FOLDER'], 'source.xlsx'),os.path.join(app.config['UPLOAD_FOLDER'], 'converted.xlsx'))
        return send_from_directory(app.config['UPLOAD_FOLDER'], convertedfile, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found"}), 404


@app.route('/convert-excel', methods=['POST'])
def convert_excel():
    # 检查请求中是否包含文件
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    
    # 如果用户没有选择文件
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    if file:
        # 保存文件到指定目录
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], sourcefile)
        file.save(filepath)

    try:
        from converter import convert
        convert(os.path.join(app.config['UPLOAD_FOLDER'], 'source.xlsx'),os.path.join(app.config['UPLOAD_FOLDER'], 'converted.xlsx'))
        return send_from_directory(app.config['UPLOAD_FOLDER'], convertedfile, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found"}), 404

if __name__ == '__main__':
    app.run(host='0.0.0.0',port=10999,debug=True)
