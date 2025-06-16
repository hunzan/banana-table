from flask import Flask, render_template, request, send_file, jsonify, make_response
import re
from generate_docx import generate_schedule_docx
from load_csv import parse_calendar_csv
from download_csv import generate_csv
from calendar_docx import generate_monthly_calendar_docx
from calendar_download_csv import generate_calendar_csv
from load_csv import load_custom_csv
from custom_docx import generate_custom_docx
from urllib.parse import quote

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/schedule')
def schedule_form():
    return render_template('schedule_form.html')

@app.route('/howto')
def howto():
    return render_template('howto.html')

@app.route('/calendar')
def calendar_form():
    return render_template('calendar_form.html')

@app.route('/howto_calendar')
def howto_calendar():
    return render_template('howto_calendar.html')

@app.route("/custom")
def custom_form():
    return render_template("custom_form.html")

@app.route('/howto_custom')
def howto_custom():
    return render_template('howto_custom.html')

# ----- 匯入功能區 -----
@app.route('/upload-json', methods=['POST'])
def upload_json():
    try:
        data = request.get_json()
        print("🔍 收到 JSON：", data)
        return jsonify(data)
    except Exception as e:
        print(f"❌ 錯誤：{e}")
        return 'Failed to parse JSON', 400

@app.route('/upload-calendar-csv', methods=['POST'])
def upload_calendar_csv():
    file = request.files.get('csv')
    if not file:
        return "No file uploaded", 400

    try:
        data = parse_calendar_csv(file)
        return jsonify(data)
    except Exception as e:
        print("❌ 解析月曆 CSV 失敗：", e)
        return "CSV format error", 400

@app.route("/upload-custom-csv", methods=["POST"])
def upload_custom_csv():
    return load_custom_csv()

# ----- 匯出功能區 -----
@app.route('/download-docx', methods=['POST'])
def download_docx():
    data = request.get_json()
    doc_io = generate_schedule_docx(data)

    # 從使用者填的 title 欄位取得檔名，並移除不合法字元
    title = data.get("title", "").strip()
    safe_title = re.sub(r'[\\/:*?"<>|]', '_', title) or "課表"

    return send_file(
        doc_io,
        as_attachment=True,
        download_name=f"{safe_title}.docx"
    )

@app.route('/download-csv', methods=['POST'])
def download_csv_route():
    data = request.get_json()
    return generate_csv(data)

@app.route('/download-month-docx', methods=['POST'])
def download_month_docx():
    data = request.get_json()

    doc_io = generate_monthly_calendar_docx(data)

    title = data.get("title", "").strip()
    safe_title = re.sub(r'[\\/:*?"<>|]', '_', title) or "月行事曆"

    return send_file(
        doc_io,
        as_attachment=True,
        download_name=f"{safe_title}.docx"
    )

@app.route("/download-calendar-csv", methods=["POST"])
def download_calendar_csv():
    data = request.json
    return generate_calendar_csv(data) 

@app.route("/generate-custom-docx", methods=["POST"])
def generate_custom_docx_route():
    titles = request.form.getlist("titles")
    contents = request.form.getlist("contents")
    table_title = request.form.get("tableTitle", "自訂表格")

    font_name = request.form.get("font", "DFKai-SB")
    layout = request.form.get("layout", "portrait")
    border_width = request.form.get("borderWidth", "8")
    
    docx_file = generate_custom_docx(
        titles, contents, table_title, border_width, font_name, layout
    )

    filename = f"{table_title}.docx"  # ← 用輸入的表格標題命名檔案
    return send_file(
        docx_file,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

from urllib.parse import quote

@app.route("/download-custom-csv", methods=["POST"])
def download_custom_csv():
    table_title = request.form.get("tableTitle", "自訂表格")  # 使用者輸入的標題
    titles = request.form.getlist("titles")
    contents = request.form.getlist("contents")

    lines = [f'"{t}","{c}"' for t, c in zip(titles, contents)]
    csv_data = "\n".join(lines)

    # 將檔名處理為 URL encoded（避免中文引發 UnicodeEncodeError）
    safe_filename = quote(f"{table_title}.csv")

    response = make_response(csv_data)
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{safe_filename}"
    response.mimetype = "text/csv"
    return response

if __name__ == '__main__':
    app.run(debug=False, host="0.0.0.0", port=5000)

