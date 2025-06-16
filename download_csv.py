import io
import csv
from flask import send_file

def generate_csv(data):
    title = data.get('title', '')
    weeks = data.get('weeks', [])
    periods = data.get('periods', [])
    times = data.get('times', [])
    remarks = data.get('remarks', '')  # 說明：表格上方文字
    note = data.get('note', '')        # 備註：表格下方文字
    breaks = data.get('breaks', [])
    content = data.get('content', [])

    output = io.StringIO()
    writer = csv.writer(output)

    # 用機器讀得懂的欄位格式
    writer.writerow(['[title]', title])
    writer.writerow(['[remarks]', remarks])
    writer.writerow(['[weeks]'] + weeks)
    writer.writerow(['[periods]'] + periods)
    if any(times):
        writer.writerow(['[times]'] + times)
    for b in breaks:
        writer.writerow(['[breaks]', b])
    for row in content:
        writer.writerow(['[content]'] + row)
    writer.writerow(['[note]', note])

    output.seek(0)

    filename = f"{title.strip()}.csv" if title.strip() else "蕉表格_備份.csv"

    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name=filename
    )
