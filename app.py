"""
Flask web application: TT document generator.
"""
import os
import json
from flask import Flask, render_template, request, jsonify, send_file
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

from parser import parse_contract_file, parse_excel_input
from generator import generate_tt_excel

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

MONTHS = ['JAN.', 'FEB.', 'MAR.', 'APR.', 'MAY', 'JUN.',
          'JUL.', 'AUG.', 'SEP.', 'OCT.', 'NOV.', 'DEC.']


def today_invoice_date():
    now = datetime.now(ZoneInfo('Asia/Shanghai'))
    return f"{MONTHS[now.month - 1]} {now.day:02d}, {now.year}"


@app.errorhandler(500)
def internal_error(e):
    import traceback
    return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': '文件超过 32MB 限制'}), 413


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/parse', methods=['POST'])
def parse_files():
    """
    Accept multipart form with optional 'contract' and 'excel' files.
    Return JSON with pre-parsed fields.
    """
    result = {'contract': {}, 'excel': {'type': 'unknown', 'grades': []}}

    contract_file = request.files.get('contract')
    if contract_file and contract_file.filename:
        try:
            result['contract'] = parse_contract_file(
                contract_file.read(), contract_file.filename)
        except Exception as e:
            result['contract_error'] = str(e)

    excel_file = request.files.get('excel')
    if excel_file and excel_file.filename:
        try:
            result['excel'] = parse_excel_input(
                excel_file.read(), excel_file.filename)
        except Exception as e:
            result['excel_error'] = str(e)

    return jsonify(result)


@app.route('/generate', methods=['POST'])
def generate():
    """
    Accept JSON body with all TT fields, return Excel file download.
    """
    data = request.get_json(force=True)
    if not data:
        return jsonify({'error': 'No data provided'}), 400
    data['invoice_date'] = today_invoice_date()

    try:
        xlsx_bytes = generate_tt_excel(data)
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

    contract_no = data.get('contract_no', 'TT')
    filename    = f"{contract_no} TT文件.xlsx"

    return send_file(
        BytesIO(xlsx_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
