"""
Report 1.7 - Login Logout DAY
Columns: Login Time(str), Logout Time(str), Logout Reason(str), Total Login Duration(Td)
"""
import pandas as pd
from .utils import read_cuic_xls, timedelta_to_hms, df_to_records, build_excel_response


def process(file_data: bytes, filename: str = '') -> dict:
    df = read_cuic_xls(file_data)

    col_map = {}
    for c in df.columns:
        cl = c.lower().strip()
        if 'login time' in cl:
            col_map[c] = 'Login Time'
        elif 'logout time' in cl:
            col_map[c] = 'Logout Time'
        elif 'reason' in cl:
            col_map[c] = 'Logout Reason'
        elif 'duration' in cl or 'total' in cl:
            col_map[c] = 'Duration'
    df = df.rename(columns=col_map)

    if 'Login Time' in df.columns:
        df = df[df['Login Time'].notna() & (df['Login Time'].astype(str).str.strip() != '')].reset_index(drop=True)

    if 'Duration' in df.columns:
        df['Duration'] = df['Duration'].apply(timedelta_to_hms)

    def edate(v):
        try:
            return pd.to_datetime(str(v).strip()).strftime('%d/%m/%Y')
        except Exception:
            return str(v)[:10]

    if 'Login Time' in df.columns:
        df['Login Date'] = df['Login Time'].apply(edate)

    reason_counts = {}
    if 'Logout Reason' in df.columns:
        reason_counts = df['Logout Reason'].astype(str).str.strip().value_counts().to_dict()

    labels = df['Login Date'].tolist() if 'Login Date' in df.columns else [str(i+1) for i in range(len(df))]
    out = [c for c in ['Login Date','Login Time','Logout Time','Logout Reason','Duration'] if c in df.columns]

    return {
        'report': '1_7',
        'title': '1.7 Login/Logout DAY - รายงานการเข้าออกงาน',
        'labels': labels,
        'datasets': {
            'reason_labels': list(reason_counts.keys()),
            'reason_counts': list(reason_counts.values()),
            'durations': df['Duration'].tolist() if 'Duration' in df.columns else [],
        },
        'table': df_to_records(df[out]), 'columns': out,
    }


def export_excel(payload: dict) -> bytes:
    data = payload.get('data', payload)
    records = data.get('table', [])
    cols = data.get('columns', list(records[0].keys()) if records else [])
    df = pd.DataFrame(records, columns=cols)
    return build_excel_response(df, sheet_name='Login Logout', title='1.7 Login Logout DAY')
