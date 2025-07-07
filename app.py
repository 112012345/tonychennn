import os
from flask import Flask, request, send_file, render_template
from 自動化小工具 import run_batch_process

app = Flask(__name__)

# 確保資料夾存在
os.makedirs('uploads/word', exist_ok=True)
os.makedirs('uploads', exist_ok=True)
os.makedirs('output', exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    word_files = request.files.getlist('word_files')
    excel_file = request.files['excel_file']

    word_dir = 'uploads/word'
    # 清空舊檔案
    for f in os.listdir(word_dir):
        os.remove(os.path.join(word_dir, f))

    for f in word_files:
        f.save(os.path.join(word_dir, f.filename))

    excel_path = os.path.join('uploads', excel_file.filename)
    excel_file.save(excel_path)

    # 執行自動化處理
    run_batch_process(word_dir, excel_path, '差異及', '文件編號')

    # 回傳處理後的 Excel
    processed_excel = 'output/03_SSR差異分析方案評分工作表_授信.xlsx'
    return send_file(processed_excel, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
