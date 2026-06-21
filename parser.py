"""
Parse contract PDF and input Excel (ERP / 明细 / DPL出库通知单) into a structured dict.
"""
import re
import io
import math
import os
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET


# ── OCR fallback ────────────────────────────────────────────────────────────
def _ocr_pdf(file_bytes: bytes) -> str:
    """OCR only the first page (grayscale, 2x) to stay within 512MB memory limit.
    All key contract fields (number, buyer, commodity, price) are on page 1.
    Grayscale cuts memory to 1/3 of RGB; single page avoids OOM on free-tier hosts.
    """
    try:
        import fitz
        import pytesseract
        from PIL import Image

        doc = fitz.open(stream=file_bytes, filetype='pdf')
        if doc.page_count == 0:
            return ''
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2), colorspace=fitz.csGRAY)
        img = Image.frombytes('L', [pix.width, pix.height], pix.samples)
        return pytesseract.image_to_string(img, lang='eng')
    except Exception:
        return ''


# ── Contract text parsing ───────────────────────────────────────────────────
def _parse_contract_text(full_text: str) -> dict:
    result = {}
    full_text = (full_text or '').replace('\xa0', ' ')
    full_text = re.sub(r'[ \t]+', ' ', full_text)
    if not full_text.strip():
        return result

    # Contract number  (e.g. 26YCR017, 26YWR001 — may start with digits)
    m = re.search(r'CONTRACT\s+NO\.?:?\s*([A-Z0-9]{5,})', full_text, re.I)
    if m:
        result['contract_no'] = m.group(1).strip()
        result['invoice_no']  = result['contract_no'] + 'V'

    # Buyer name
    m = re.search(r'BUYER\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['buyer_name'] = m.group(1).strip()

    # Buyer address
    m = re.search(r'BUYER\s*:[^\n]+\n\s*ADD\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['buyer_address'] = m.group(1).strip()

    # NIT / RUC / FAX ID (buyer tax/reference number)
    m = re.search(r'((?:NIT|RUC|FAX\s*ID|TAX\s*ID|VAT)\s*:\s*[\d\-]+)', full_text, re.I)
    if m:
        result['buyer_ref'] = m.group(1).strip()

    # Commodity
    m = re.search(r'COMMODITY\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['commodity'] = m.group(1).strip()

    # Unit price from table row (e.g. "1  1.20X1219  300  588  176,400.00")
    m = re.search(
        r'\d+\s+[\d.X*]+\s+[\d,]+\s+([\d,]+(?:\.\d+)?)\s+[\d,]+(?:\.\d+)?',
        full_text)
    if m:
        try:
            result['unit_price'] = float(m.group(1).replace(',', ''))
        except ValueError:
            pass

    # Grade
    composite_grades = re.findall(
        r'\b([A-Z]{2,}\d*[A-Z]*(?:-[A-Z0-9]+)?/[A-Z]{2,}\d*[A-Z]*(?:-[A-Z0-9]+)?)\b',
        full_text, re.I)
    m = next((g for g in composite_grades if '-' in g or any(ch.isdigit() for ch in g)), None)
    grade_pat = re.compile(
        r'\b(SPCC[-\w]*|SAE\d+[A-Z]?|S\d{3}[A-Z]{1,3}|Q\d+[A-Z]?|[A-Z]{2,}\d+[-\w]*)\b')
    if m:
        result['detected_grade'] = m.strip().upper()
    else:
        grades_found = grade_pat.findall(full_text)
        if grades_found:
            result['detected_grade'] = grades_found[0]

    # Contract total quantity (used for prepayment allocation)
    m = re.search(r'TOTAL\s+QUANTITY\s*:?\s*([\d,]+(?:\.\d+)?)\s*MT', full_text, re.I)
    if m:
        try:
            result['contract_total_quantity'] = float(m.group(1).replace(',', ''))
        except ValueError:
            pass

    # Shipping mark
    m = re.search(r'SHIPPING\s+MARK\s*:?\s*([^\n]+)', full_text, re.I)
    if m:
        result['shipping_mark'] = m.group(1).strip()

    # Price terms
    m = re.search(r'PRICE\s+TERMS\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['price_terms'] = m.group(1).strip().rstrip('.')

    result['country_of_origin'] = 'CHINA'
    return result


# ── PDF / Word parsing ─────────────────────────────────────────────────────
def parse_contract_pdf(file_bytes: bytes) -> dict:
    try:
        import fitz
        doc = fitz.open(stream=file_bytes, filetype='pdf')
        full_text = '\n'.join(page.get_text() for page in doc)
    except Exception:
        return {}

    if not full_text.strip():
        full_text = _ocr_pdf(file_bytes)   # scanned / image PDF → OCR
    return _parse_contract_text(full_text)


def parse_contract_word(file_bytes: bytes, filename: str = '') -> dict:
    """Parse Word contracts without OCR. Supports .docx directly and .doc via antiword."""
    fname = (filename or '').lower()
    text = ''

    if fname.endswith('.docx'):
        try:
            parts = []
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                xml_bytes = zf.read('word/document.xml')
            root = ET.fromstring(xml_bytes)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            for para in root.iterfind('.//w:p', ns):
                line = ''.join(t.text or '' for t in para.iterfind('.//w:t', ns)).strip()
                if line:
                    parts.append(line)
            text = '\n'.join(parts)
        except Exception:
            text = ''
    elif fname.endswith('.doc'):
        # Binary .doc files need an external reader. Docker installs antiword.
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.doc') as tmp:
                tmp.write(file_bytes)
                tmp_path = tmp.name
            try:
                proc = subprocess.run(
                    ['antiword', tmp_path],
                    check=False,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.DEVNULL,
                    timeout=20,
                )
                text = proc.stdout.decode('utf-8', errors='ignore')
            finally:
                os.unlink(tmp_path)
        except Exception:
            text = ''

    return _parse_contract_text(text)


def parse_contract_file(file_bytes: bytes, filename: str) -> dict:
    fname = (filename or '').lower()
    if fname.endswith(('.doc', '.docx')):
        return parse_contract_word(file_bytes, filename)
    return parse_contract_pdf(file_bytes)


# ── Excel helpers ────────────────────────────────────────────────────────────
def _format_size(raw: str) -> str:
    """'1.2*1219' → '1.20X1219',  '3.28*1286' → '3.28X1286',  '5.5' → '5.5'"""
    s = str(raw).strip()
    if '*' in s:
        parts = s.split('*', 1)
        try:
            thickness = float(parts[0])
            return f'{thickness:.2f}X{parts[1]}'
        except ValueError:
            return s.replace('*', 'X')
    return s


def _is_erp(df) -> bool:
    cols = [str(c) for c in df.columns]
    return (any('牌号' in c for c in cols) and
            any('重量' in c for c in cols))


def _is_dpl(df) -> bool:
    """出库通知单 / 集港明细：前几行含 '出库通知单' 或列头含 '材质'+'毛重'+'卷号'"""
    # Check first 10 rows for DPL title keywords
    for i in range(min(10, len(df))):
        row_str = ' '.join(str(v) for v in df.iloc[i])
        if '出库通知单' in row_str or '集港明细' in row_str or 'PACKING LIST' in row_str:
            return True
    # Also check column headers row for characteristic columns
    for i in range(min(10, len(df))):
        vals = [str(v) for v in df.iloc[i]]
        if ('材质' in vals or 'QUAL.' in vals) and ('毛重' in vals or '卷号' in vals):
            return True
    return False


# ── Main entry point ────────────────────────────────────────────────────────
def parse_excel_input(file_bytes: bytes, filename: str) -> dict:
    """
    Detect format and parse.  Returns:
      {
        'type': 'erp' | 'detail' | 'dpl' | 'unknown',
        'contract_no': str (DPL only),
        'grades': [ {grade, size, quantity_mt, num_coils, gross_weight_mt, unit_price}, ... ]
      }
    """
    try:
        import pandas as pd
    except ImportError:
        return {'type': 'unknown', 'grades': []}

    fname = filename.lower()
    try:
        if fname.endswith('.xlsx'):
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=0, engine='openpyxl')
        else:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')
    except Exception as e:
        return {'type': 'unknown', 'grades': [], 'error': str(e)}

    if fname.endswith('.xlsx') and _is_erp(df_raw):
        return _parse_erp(df_raw)

    if _is_dpl(df_raw):
        return _parse_dpl(df_raw)

    return _parse_detail(df_raw)


# ── ERP format ───────────────────────────────────────────────────────────────
def _parse_erp(df) -> dict:
    cols       = list(df.columns)
    grade_col  = next((c for c in cols if '牌号'  in str(c)), None)
    weight_col = next((c for c in cols if '重量'  in str(c)), None)
    thick_col  = next((c for c in cols if '厚'    in str(c)), None)

    if grade_col is None or weight_col is None:
        return {'type': 'unknown', 'grades': []}

    df = df.dropna(subset=[grade_col])
    df = df[df[grade_col].astype(str).str.strip() != '']

    grades_out = []
    for grade, grp in df.groupby(grade_col, sort=False):
        size_val = ''
        if thick_col:
            v = str(grp[thick_col].iloc[0]).strip()
            if v not in ('', '0', '0.0', 'nan'):
                size_val = _format_size(v)

        net = round(float(grp[weight_col].sum()), 3)
        grades_out.append({
            'grade':           str(grade).strip(),
            'size':            size_val,
            'quantity_mt':     net,
            'num_coils':       len(grp),
            'gross_weight_mt': net,
            'unit_price':      None,
        })

    return {'type': 'erp', 'grades': grades_out}


# ── DPL / 出库通知单 format ──────────────────────────────────────────────────
def _parse_dpl(df) -> dict:
    """
    Layout:
      Row 0-2 : title rows
      Row 3   : 外销合同号 (col 3/4 label, col 5 value)
      Row 7   : Chinese column headers  → 规格 材质 件数 毛重 总重量 卷号 ...
      Row 8   : English column headers  → SIZE QUAL. PCS WEIGHT COIL NO ...
      Row 9+  : data rows (skip '小计' / '合计' rows)
    """
    result = {'type': 'dpl', 'grades': []}

    # ── Extract contract number from header ──────────────────────────────
    for i in range(min(8, len(df))):
        row = list(df.iloc[i])
        for j, v in enumerate(row):
            if '外销合同号' in str(v) and j + 1 < len(row):
                for k in range(j + 1, min(j + 6, len(row))):
                    val = str(row[k]).strip()
                    if val and val not in ('nan', '外销合同号'):
                        result['contract_no'] = val
                        break

    # ── Find the data header row (contains '规格' and '材质') ─────────────
    header_row = None
    for i in range(min(12, len(df))):
        vals = [str(v) for v in df.iloc[i]]
        if '规格' in vals and ('材质' in vals or 'QUAL.' in vals):
            header_row = i
            break

    if header_row is None:
        return result

    headers = [str(v).strip() for v in df.iloc[header_row]]

    def col(keywords):
        for kw in keywords:
            for j, h in enumerate(headers):
                if kw in h:
                    return j
        return None

    size_col   = col(['规格', 'SIZE'])
    grade_col  = col(['材质', 'QUAL'])
    gross_col  = col(['毛重', 'WEIGHT'])
    coil_col   = col(['卷号', 'COIL NO'])

    if size_col is None or gross_col is None:
        return result

    # ── Parse data rows ──────────────────────────────────────────────────
    groups: dict = {}   # key: (size, grade)

    for i in range(header_row + 2, len(df)):   # skip English header row too
        row = list(df.iloc[i])

        # Skip subtotal / total rows
        row_str = ' '.join(str(v) for v in row)
        if '小计' in row_str or '合计' in row_str:
            continue

        # Must have a numeric gross weight
        try:
            gw = float(row[gross_col])
            if math.isnan(gw) or gw <= 0:
                continue
        except (TypeError, ValueError, IndexError):
            continue

        size  = _format_size(str(row[size_col]).strip()) if size_col is not None else ''
        grade = str(row[grade_col]).strip() if grade_col is not None else ''
        if grade in ('nan', ''):
            grade = ''

        key = (size, grade)
        if key not in groups:
            groups[key] = {'coils': 0, 'gross': 0.0}
        groups[key]['coils'] += 1
        groups[key]['gross']  = round(groups[key]['gross'] + gw, 3)

    for (size, grade), g in groups.items():
        gross = round(g['gross'], 3)
        result['grades'].append({
            'grade':           grade,
            'size':            size,
            'quantity_mt':     gross,   # net = gross (no separate 净重 column)
            'num_coils':       g['coils'],
            'gross_weight_mt': gross,
            'unit_price':      None,
        })

    return result


# ── 明细 format ──────────────────────────────────────────────────────────────
def _parse_detail(df) -> dict:
    # Find header row
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v) for v in row]
        if any('净重' in v or '卷号' in v or '规格' in v for v in vals):
            header_row = i
            break

    if header_row is None:
        net_col, gross_col, spec_col = 5, 6, 3
        start_row = 2
    else:
        cols = list(df.iloc[header_row])
        def find_col(keywords):
            for j, c in enumerate(cols):
                if any(k in str(c) for k in keywords):
                    return j
            return None
        net_col   = find_col(['净重'])
        gross_col = find_col(['毛重'])
        spec_col  = find_col(['规格'])
        grade_col = find_col(['牌号', '材质', '钢种', 'GRADE', 'QUAL'])
        start_row = header_row + 1

    data_rows = df.iloc[start_row:]

    def is_data_row(row):
        try:
            v = row.iloc[net_col] if net_col is not None else None
            if v is None:
                return False
            float(v)
            return True
        except (TypeError, ValueError):
            return False

    data_rows = data_rows[data_rows.apply(is_data_row, axis=1)]
    if data_rows.empty:
        return {'type': 'detail', 'grades': []}

    groups: dict = {}
    for _, row in data_rows.iterrows():
        spec = str(row.iloc[spec_col]).strip() if spec_col is not None else ''
        grade = str(row.iloc[grade_col]).strip() if 'grade_col' in locals() and grade_col is not None else ''
        if grade.lower() == 'nan':
            grade = ''
        key = (spec, grade)
        if key not in groups:
            groups[key] = {'coils': 0, 'net': 0.0, 'gross': 0.0}
        groups[key]['coils'] += 1
        try:
            groups[key]['net'] = round(groups[key]['net'] + float(row.iloc[net_col]), 3)
        except (TypeError, ValueError):
            pass
        if gross_col is not None:
            try:
                groups[key]['gross'] = round(
                    groups[key]['gross'] + float(row.iloc[gross_col]), 3)
            except (TypeError, ValueError):
                pass

    grades_out = []
    for (spec, grade), g in groups.items():
        net = round(g['net'], 3)
        if math.isnan(net) or net == 0:
            continue
        gross = round(g['gross'], 3) if g['gross'] else net
        grades_out.append({
            'grade':           grade,
            'size':            _format_size(spec),
            'quantity_mt':     net,
            'num_coils':       g['coils'],
            'gross_weight_mt': gross,
            'unit_price':      None,
        })

    return {'type': 'detail', 'grades': grades_out}
