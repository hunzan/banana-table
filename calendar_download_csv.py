# calendar_download_csv.py
import io
import csv
from flask import send_file
from datetime import datetime

def generate_calendar_csv(data):
    font = data.get('font', '')
    layout = data.get('layout', '')
    title = data.get('title', '月行事曆')
    year = data.get('year', '')
    month = data.get('month', '')
    fixed_events = data.get('fixedEvents', [])  # list of strings
    specific_events = data.get('specificEvents', [])  # list of dicts

    output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow(['[font]', font])
    writer.writerow(['[layout]', layout])
    writer.writerow(['[title]', title])
    writer.writerow(['[year]', year])
    writer.writerow(['[month]', month])

    for item in fixed_events:
        writer.writerow(['[fixed]', item])

    for event in specific_events:
        date_str = event.get('date', '')
        try:
            # 若為 YYYY-MM-DD 格式，轉為 MM/DD
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%m/%d")
        except Exception:
            # 若格式錯誤，直接保留原樣
            formatted_date = date_str

        writer.writerow(['[event]', formatted_date, event.get('time', ''), event.get('task', '')])

    output.seek(0)

    filename = f"{title.strip()}_備份.csv" if title.strip() else "月曆表格_備份.csv"

    return send_file(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name=filename
    )
