"""
Report 1.10 - Voicemail
รายงานการฝากข้อความเสียงจากลูกค้า (ดึงจาก CAR - CSV format)
Hunt Pilot 7100 summary by hour / by day
"""
import pandas as pd
import io
from .utils import safe_int, df_to_records, build_excel_response


def process(file_data: bytes, filename: str = '') -> dict:
    # CAR exports CSV files
    df = None
    errors = []

    try:
        df = pd.read_csv(io.BytesIO(file_data))
    except Exception as e1:
        errors.append(str(e1))
        try:
            df = pd.read_csv(io.BytesIO(file_data), encoding='utf-8-sig')
        except Exception as e2:
            errors.append(str(e2))
            try:
                # Try as Excel
                import pandas as pd
                df = pd.read_excel(io.BytesIO(file_data))
            except Exception as e3:
                errors.append(str(e3))
                raise ValueError(f"ไม่สามารถอ่านไฟล์ได้: {'; '.join(errors)}")

    # Determine if this is Hour or Day report
    is_hour = 'hour' in filename.lower() if filename else False

    col_map = {}
    for c in df.columns:
        cl = c.lower().strip()
        if 'hour' in cl or 'time' in cl or 'day' in cl or 'date' in cl:
            col_map[c] = 'Time'
        elif 'call' in cl or 'count' in cl or 'total' in cl or 'volume' in cl:
            col_map[c] = 'Count'
    df = df.rename(columns=col_map)

    if 'Count' not in df.columns and len(df.columns) >= 2:
        # Assume last numeric column is count
        numeric_cols = df.select_dtypes(include='number').columns.tolist()
        if numeric_cols:
            df = df.rename(columns={numeric_cols[-1]: 'Count'})

    if 'Count' in df.columns:
        df['Count'] = df['Count'].apply(safe_int)

    labels = []
    if 'Time' in df.columns:
        labels = [str(v) for v in df['Time']]
    else:
        labels = [str(i) for i in range(len(df))]

    counts = df['Count'].tolist() if 'Count' in df.columns else [0] * len(df)
    total = sum([safe_int(c) for c in counts])

    show_cols = [c for c in ['Time', 'Count'] if c in df.columns]
    if not show_cols:
        show_cols = df.columns.tolist()[:3]

    return {
        'report': '1_10',
        'title': '1.10 รายงานการฝากข้อความเสียง (Voicemail)',
        'is_hour': is_hour,
        'labels': labels,
        'datasets': {'counts': [safe_int(c) for c in counts]},
        'total': total,
        'table': df_to_records(df[show_cols]),
        'columns': show_cols,
    }


def export_excel(payload: dict) -> bytes:
    data = payload.get('data', payload)
    records = data.get('table', [])
    cols = data.get('columns', list(records[0].keys()) if records else [])
    df = pd.DataFrame(records, columns=cols)
    return build_excel_response(df, sheet_name='Voicemail',
                                title='1.10 รายงานการฝากข้อความเสียง')
