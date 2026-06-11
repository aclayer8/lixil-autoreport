from flask import Flask, render_template, request, jsonify, send_file
import io, os, re, traceback, webbrowser, threading, time
from config import APP_HOST, APP_PORT, AUTO_OPEN_BROWSER, SHOW_ERROR_TRACE, MAX_UPLOAD_MB

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_MB * 1024 * 1024

# Import all report processors
from modules.report_csat import process as pcsat, export_excel as excsat
from modules.report_1_1 import process as p1_1, export_excel as ex1_1
from modules.report_1_2 import process as p1_2, export_excel as ex1_2
from modules.report_1_3 import process as p1_3, export_excel as ex1_3
from modules.report_1_4 import process as p1_4, export_excel as ex1_4
from modules.report_1_5 import process as p1_5, export_excel as ex1_5
from modules.report_1_6 import process as p1_6, export_excel as ex1_6
from modules.report_1_7 import process as p1_7, export_excel as ex1_7
from modules.report_1_8 import process as p1_8, export_excel as ex1_8
from modules.report_1_9 import process as p1_9, export_excel as ex1_9
from modules.report_1_10 import process as p1_10, export_excel as ex1_10
from modules.report_1_11 import process as p1_11, export_excel as ex1_11

PROCESSORS = {
    'csat': pcsat, '1_1': p1_1, '1_2': p1_2, '1_3': p1_3, '1_4': p1_4,
    '1_5': p1_5, '1_6': p1_6, '1_7': p1_7, '1_8': p1_8,
    '1_9': p1_9, '1_10': p1_10, '1_11': p1_11
}
EXPORTERS = {
    'csat': excsat, '1_1': ex1_1, '1_2': ex1_2, '1_3': ex1_3, '1_4': ex1_4,
    '1_5': ex1_5, '1_6': ex1_6, '1_7': ex1_7, '1_8': ex1_8,
    '1_9': ex1_9, '1_10': ex1_10, '1_11': ex1_11
}

REPORT_NAMES = {
    'csat': 'CSAT_Report', '1_1': 'CSQ_Day', '1_2': 'CSQ_Hour', '1_3': 'Agent_States_Day',
    '1_4': 'All_Agents_States_Month', '1_5': 'Agent_States_YTD',
    '1_6': 'Not_Ready_Count', '1_7': 'Login_Logout',
    '1_8': 'Audio_Text_Maintenance', '1_9': 'Audio_Text_Problem',
    '1_10': 'Voicemail', '1_11': 'Reserved_Talking_Ratio'
}

def converted_download_name(report_id, payload):
    data = payload.get('data') if isinstance(payload, dict) else {}
    source = payload.get('sourceFilename') or (data or {}).get('sourceFilename') or ''
    source = os.path.basename(str(source)).strip()
    if not source:
        source = REPORT_NAMES.get(report_id, f'report_{report_id}')
    stem = re.sub(r'\.[^.]+$', '', source)
    safe_stem = re.sub(r'[\\/:*?"<>|]+', '_', stem).strip(' ._') or REPORT_NAMES.get(report_id, f'report_{report_id}')
    return f'converted_{safe_stem}.xlsx'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process/<report_id>', methods=['POST'])
def process_report(report_id):
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No uploaded file found'}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'}), 400
    if report_id not in PROCESSORS:
        return jsonify({'success': False, 'error': f'Unknown report: {report_id}'}), 400
    try:
        data = f.read()
        result = PROCESSORS[report_id](data, f.filename)
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        payload = {'success': False, 'error': str(e)}
        if SHOW_ERROR_TRACE:
            payload['trace'] = traceback.format_exc()
        return jsonify(payload), 500

@app.route('/api/export/<report_id>/excel', methods=['POST'])
def export_excel(report_id):
    if report_id not in EXPORTERS:
        return jsonify({'error': 'Unknown report'}), 400
    try:
        payload = request.json or {}
        excel_bytes = EXPORTERS[report_id](payload)
        return send_file(
            io.BytesIO(excel_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=converted_download_name(report_id, payload)
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/export/<report_id>/pdf', methods=['POST'])
def export_pdf_route(report_id):
    # PDF generation handled client-side via window.print()
    return jsonify({'message': 'use client-side print'})

def open_browser():
    time.sleep(1.5)
    webbrowser.open(f'http://127.0.0.1:{APP_PORT}')

if __name__ == '__main__':
    if AUTO_OPEN_BROWSER:
        threading.Thread(target=open_browser, daemon=True).start()
    print("\n" + "="*50)
    print("  LIXIL Report App")
    print(f"  http://{APP_HOST}:{APP_PORT}")
    print("  Press Ctrl+C to stop")
    print("="*50 + "\n")
    app.run(debug=False, port=APP_PORT, host=APP_HOST, use_reloader=False)


