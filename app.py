import os
import uuid
import traceback
import threading
import time
import shutil
from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for
from functools import wraps
from config import UPLOAD_FOLDER, OUTPUT_FOLDER
from parsers.xlsx_parser import parse_xlsx
from engine.classifier import classify_data
from engine.note_manager import build_note_map
from engine.workbook_builder import build_icai_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.secret_key = os.environ.get('SECRET_KEY', 'asar-icai-nce-2026-key')

APP_PASSWORD = os.environ.get('APP_PASSWORD', '123456')

jobs = {}
MAX_FILE_AGE_HOURS = 24


def _cleanup_old_files():
    while True:
        time.sleep(3600)
        cutoff = time.time() - MAX_FILE_AGE_HOURS * 3600
        for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
            if not os.path.exists(folder):
                continue
            for name in os.listdir(folder):
                path = os.path.join(folder, name)
                try:
                    if os.path.getmtime(path) < cutoff:
                        if os.path.isdir(path):
                            shutil.rmtree(path, ignore_errors=True)
                        else:
                            os.remove(path)
                except OSError:
                    pass


threading.Thread(target=_cleanup_old_files, daemon=True).start()


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        pwd = request.form.get('password', '')
        if pwd == APP_PASSWORD:
            session['logged_in'] = True
            session.permanent = True
            return redirect(url_for('index'))
        error = 'Galat password. Phir se try karo.'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
@login_required
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
@login_required
def convert():
    try:
        job_id = str(uuid.uuid4())[:8]
        job_dir = os.path.join(UPLOAD_FOLDER, job_id)
        os.makedirs(job_dir, exist_ok=True)

        cy_file = request.files.get('cy_file')
        py_file = request.files.get('py_file')

        if not cy_file or not py_file:
            return jsonify({'error': 'Dono files upload karo — Current Year aur Previous Year'}), 400

        cy_ext = cy_file.filename.rsplit('.', 1)[-1].lower() if '.' in cy_file.filename else ''
        py_ext = py_file.filename.rsplit('.', 1)[-1].lower() if '.' in py_file.filename else ''

        if cy_ext not in ('xlsx', 'xls'):
            return jsonify({'error': f'Current Year file format ({cy_ext}) abhi supported nahi hai. Sirf .xlsx upload karo.'}), 400
        if py_ext not in ('xlsx', 'xls'):
            return jsonify({'error': f'Previous Year file format ({py_ext}) abhi supported nahi hai. Sirf .xlsx upload karo.'}), 400

        cy_path = os.path.join(job_dir, 'cy_' + cy_file.filename)
        py_path = os.path.join(job_dir, 'py_' + py_file.filename)
        cy_file.save(cy_path)
        py_file.save(py_path)

        cy_parsed_list = parse_xlsx(cy_path)
        py_parsed_list = parse_xlsx(py_path)

        if not cy_parsed_list:
            return jsonify({'error': 'Current Year file mein koi financial data nahi mila. T-format (Trading + P&L + Balance Sheet) expected hai.'}), 400
        if not py_parsed_list:
            return jsonify({'error': 'Previous Year file mein koi financial data nahi mila. T-format (Trading + P&L + Balance Sheet) expected hai.'}), 400

        cy_parsed = cy_parsed_list[0]
        py_parsed = py_parsed_list[0]

        classified = classify_data(cy_parsed, py_parsed)
        note_info = build_note_map(classified)

        entity_name = classified.get('entity_name', 'Entity')
        safe_name = ''.join(c if c.isalnum() or c in ' _-' else '_' for c in entity_name)[:30]
        fy = classified.get('fy_cy', 'FY')
        output_filename = f'{safe_name}_FS_{fy}_ICAI_NCE.xlsx'
        output_path = os.path.join(OUTPUT_FOLDER, f'{job_id}_{output_filename}')

        build_icai_workbook(classified, note_info, output_path)

        jobs[job_id] = {
            'output_path': output_path,
            'output_filename': output_filename,
        }

        rev_cy = sum(i['amount'] for i in classified.get('revenue_from_operations_cy', []))
        rev_py = sum(i['amount'] for i in classified.get('revenue_from_operations_py', []))
        np_cy = classified.get('net_profit_cy', 0)
        np_py = classified.get('net_profit_py', 0)

        cap_cy = classified.get('capital_cy', {}).get('closing', 0)
        cap_py = classified.get('capital_py', {}).get('closing', 0)

        def _total_assets(c):
            total = 0
            for key in ['ppe', 'gold_investments', 'security_deposit',
                        'trade_receivables', 'cash_hand', 'bank_balance',
                        'fixed_deposit', 'gst_input', 'tds_advance_tax', 'other_ca']:
                val = c.get(f'{key}_{{}}'.format('cy'), [])
                if isinstance(val, list):
                    total += sum(i.get('amount', 0) for i in val)
                elif isinstance(val, (int, float)):
                    total += val
            inv = c.get('inventories_cy', 0) or c.get('closing_stock_cy', 0)
            total += inv
            return total

        retained_notes = [n['number'] for n in note_info['retained']]
        suppressed_notes = [n['title'] for n in note_info['suppressed']]

        summary = {
            'job_id': job_id,
            'entity_name': entity_name,
            'fy_cy': classified.get('fy_cy', ''),
            'fy_py': classified.get('fy_py', ''),
            'constitution': classified.get('constitution', ''),
            'revenue_cy': round(rev_cy, 2),
            'revenue_py': round(rev_py, 2),
            'net_profit_cy': round(np_cy, 2),
            'net_profit_py': round(np_py, 2),
            'capital_cy': round(cap_cy, 2),
            'capital_py': round(cap_py, 2),
            'retained_notes': len(note_info['retained']),
            'suppressed_notes': len(note_info['suppressed']),
            'suppressed_list': suppressed_notes[:10],
            'note_range': f"1 to {note_info['last_note_number']}",
        }

        return jsonify(summary)

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/download/<job_id>')
@login_required
def download(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found — expired ya galat ID'}), 404
    return send_file(
        job['output_path'],
        as_attachment=True,
        download_name=job['output_filename'],
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    import socket
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    print(f'\n  [OK] ICAI NCE Financial Statement Converter running!')
    print(f'  >>  Open in browser:  http://{local_ip}:5001')
    print(f'  >>  Share this link with staff on same WiFi\n')
    app.run(host='0.0.0.0', port=5001, debug=False)
