import csv
import io
from datetime import datetime

def parse_uploaded_csv(file_storage):
    decoded = file_storage.read().decode('utf-8-sig')
    reader = csv.reader(io.StringIO(decoded))
    rows = list(reader)

    data = {
        'title': '',
        'remarks': '',
        'weeks': [],
        'periods': [],
        'times': [],
        'breaks': [],
        'content': [],
    }

    current_section = None

    for row in rows:
        if not row or not row[0].strip():
            continue

        key = row[0].strip().lower()

        # 確認是某個區塊的開頭
        if key in ['[title]', '[remarks]', '[weeks]', '[periods]', '[times]', '[note]']:
            current_section = key
            values = [cell.strip() for cell in row[1:]]
            value = values[0] if values else ''

            if key == '[title]':
                data['title'] = value
            elif key == '[remarks]':
                data['remarks'] = value
            elif key == '[weeks]':
                data['weeks'] = values
            elif key == '[periods]':
                data['periods'] = values
            elif key == '[times]':
                data['times'] = values

        elif key == '[breaks]':
            current_section = '[breaks]'
            values = [cell.strip() for cell in row[1:]]
            if values:
                data['breaks'].append(','.join(values))

        elif key == '[content]':
            current_section = '[content]'
            values = [cell.strip() for cell in row[1:]]
            if values:
                data['content'].append(values)

        # 如果不是區塊開頭，而是持續的 content 或 breaks
        elif current_section == '[content]':
            values = [cell.strip() for cell in row]
            if values:
                data['content'].append(values)

        elif current_section == '[breaks]':
            values = [cell.strip() for cell in row]
            if values:
                data['breaks'].append(','.join(values))

    return data

def parse_calendar_csv(file_storage):
    stream = io.StringIO(file_storage.stream.read().decode("utf-8"))
    reader = csv.reader(stream)

    font = ""
    layout = ""
    title = ""
    year = ""
    month = ""

    fixed_events = []
    specific_events = []

    for row in reader:
        if not row or not row[0].strip():
            continue

        label = row[0].strip().lower()

        if label == '[title]' and len(row) > 1:
            title = row[1].strip()
        elif label == '[year]' and len(row) > 1:
            year = row[1].strip()
        elif label == '[month]' and len(row) > 1:
            month = row[1].strip()
        elif label == '[layout]' and len(row) > 1:
            layout = row[1].strip()
        elif label == '[font]' and len(row) > 1:
            font = row[1].strip()
        elif label == '[fixed]' and len(row) > 1:
            fixed_events.append(row[1].strip())
        elif label == '[event]' and len(row) >= 4:
            raw_date = row[1].strip()
            try:
                # 先嘗試將 MM/DD 轉成 datetime 物件
                parsed_date = datetime.strptime(raw_date, "%m/%d")
                # 再補上年份成 YYYY-MM-DD
                full_date = f"{year}-{parsed_date.month:02}-{parsed_date.day:02}"
            except Exception as e:
                print(f"⚠️ 特定行程日期格式錯誤：{row} → {e}")
                full_date = raw_date  # 保留原始格式（fallback）

            specific_events.append({
                "date": full_date,
                "time": row[2].strip(),
                "task": row[3].strip()
            })

    return {
        "title": title,
        "year": year,
        "month": month,
        "layout": layout,
        "font": font,
        "fixedEvents": fixed_events,
        "specificEvents": specific_events
    }

def load_custom_csv(file_path):
    with open(file_path, encoding='utf-8') as f:
        lines = f.readlines()
    titles = lines[0].strip().split(',')
    contents = lines[1].strip().split(',')
    return titles, contents

import io

def generate_custom_csv(titles, contents):
    output = io.StringIO()
    for title, content in zip(titles, contents):
        output.write(f'"{title}","{content}"\n')
    return output.getvalue()
