"""
Report 1.8 - Audio Text Maintenance
Uses pre-computed summary in cols 17-18 (topic name, listen count).
"""
import io
import pandas as pd
import openpyxl
from .utils import df_to_records, build_excel_response


def _load(file_data: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    topics, listens = [], []
    skip_headers = {'', 'None', 'nan', 'หัวข้อ', 'จำนวนครั้งที่ถูกฟัง'}
    for row in rows:
        c17 = row[17] if len(row) > 17 else None
        c18 = row[18] if len(row) > 18 else None
        label = str(c17).strip() if c17 is not None else ''
        if label and label not in skip_headers:
            try:
                cnt = int(c18) if c18 is not None else 0
            except Exception:
                cnt = 0
            topics.append(label)
            listens.append(cnt)

    if topics:
        return topics, listens

    # Fallback: count first key pressed in col A
    from collections import Counter
    kc = Counter()
    key_names = {'1':'1. ดินเผา','2':'2. ฝารองนั่ง','3':'3. อ่างอาบน้ำ',
                 '4':'4. ตู้อาบน้ำ','5':'5. กระจกเงา','6':'6. อ่างล้างจาน',
                 '7':'7. ก๊อกน้ำ','8':'8. ฟลัชวาล์ว'}
    for row in rows[2:]:
        v = row[0]
        if v is not None:
            s = str(v).strip()
            if s:
                k = s.split()[0]
                if k.isdigit():
                    kc[k] += 1
    topics  = [key_names.get(k, f'Option {k}') for k in sorted(kc)]
    listens = [kc[k] for k in sorted(kc)]
    return topics, listens


def process(file_data: bytes, filename: str = '') -> dict:
    topics, counts = _load(file_data)
    total = sum(counts)
    tdf = pd.DataFrame({'Topic': topics, 'Count': counts,
                        '%': [round(c/total*100,2) if total>0 else 0.0 for c in counts]})
    return {
        'report': '1_8',
        'title': '1.8 Audio Text Maintenance',
        'labels': topics,
        'datasets': {'counts': counts},
        'total': total,
        'table': df_to_records(tdf), 'columns': ['Topic','Count','%'],
    }


def export_excel(payload: dict) -> bytes:
    data = payload.get('data', payload)
    records = data.get('table', [])
    cols = data.get('columns', list(records[0].keys()) if records else [])
    df = pd.DataFrame(records, columns=cols)
    return build_excel_response(df, sheet_name='Audio Maintenance', title='1.8 Audio Text Maintenance')
