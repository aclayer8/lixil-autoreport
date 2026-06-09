"""
Report 1.9 - Audio Text Problem
Counts occurrences of each Problem code (extension/IVR key).
"""
import pandas as pd
from .utils import read_cuic_xls, df_to_records, build_excel_response


def process(file_data: bytes, filename: str = '') -> dict:
    df = read_cuic_xls(file_data, header_row=1, skip_footer=0)

    col_map = {}
    for c in df.columns:
        cl = c.lower().strip()
        if 'problem' in cl:
            col_map[c] = 'Problem'
        elif 'start' in cl or 'time' in cl:
            col_map[c] = 'Start Time'
    df = df.rename(columns=col_map)

    if 'Problem' not in df.columns:
        return {'report':'1_9','title':'1.9 Audio Text Problem','labels':[],
                'datasets':{'counts':[]},'total':0,'table':[],'columns':['Problem','Count','%']}

    mask = (df['Problem'].notna() &
            ~df['Problem'].astype(str).str.match(r'^\d{1,2}-\d{4}$') &
            (df['Problem'].astype(str).str.strip() != ''))
    series = df.loc[mask, 'Problem'].astype(str).str.strip()

    counts = series.value_counts()
    labels = counts.index.tolist()
    values = counts.values.tolist()
    total  = sum(values)
    tdf = pd.DataFrame({'Problem': labels, 'Count': values,
                        '%': [round(v/total*100,2) if total>0 else 0.0 for v in values]})
    return {
        'report': '1_9',
        'title': '1.9 Audio Text Problem',
        'labels': labels,
        'datasets': {'counts': values},
        'total': total,
        'table': df_to_records(tdf), 'columns': ['Problem','Count','%'],
    }


def export_excel(payload: dict) -> bytes:
    data = payload.get('data', payload)
    records = data.get('table', [])
    cols = data.get('columns', list(records[0].keys()) if records else [])
    df = pd.DataFrame(records, columns=cols)
    return build_excel_response(df, sheet_name='Audio Problem', title='1.9 Audio Text Problem')
