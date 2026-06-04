"""
Parse contract PDF and input Excel (ERP or 明细) into a structured dict.
"""
import re
import io
from typing import Optional


# ── PDF parsing ─────────────────────────────────────────────────────────────
def parse_contract_pdf(file_bytes: bytes) -> dict:
    """
    Extract key fields from a contract PDF.
    Returns a dict with whatever could be parsed (caller fills in the rest).
    """
    result = {}
    try:
        import fitz
        doc = fitz.open(stream=file_bytes, filetype='pdf')
        full_text = '\n'.join(page.get_text() for page in doc)
    except Exception:
        return result

    if not full_text.strip():
        return result   # scanned / image PDF — nothing to extract

    # Contract number  (handles "NO.:", "NO:", "NO." etc.)
    m = re.search(r'CONTRACT\s+NO\.?:?\s*([A-Z0-9]+)', full_text, re.I)
    if m:
        result['contract_no'] = m.group(1).strip()
        result['invoice_no']  = result['contract_no'] + 'V'

    # Buyer name — line after "BUYER:" or "BUYER\s*:"
    m = re.search(r'BUYER\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['buyer_name'] = m.group(1).strip()

    # Buyer address — line after "ADD:" that follows BUYER block
    m = re.search(
        r'BUYER\s*:[^\n]+\n\s*ADD\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['buyer_address'] = m.group(1).strip()

    # NIT / RUC (buyer reference)
    m = re.search(r'((?:NIT|RUC)\s*:\s*[\d\-]+)', full_text, re.I)
    if m:
        result['buyer_ref'] = m.group(1).strip()

    # Commodity
    m = re.search(r'COMMODITY\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['commodity'] = m.group(1).strip()

    # Unit price — look for table row pattern: number  price  amount
    # e.g.  1  1.20X1219  300  588  176,400.00
    m = re.search(
        r'\d+\s+[\d.X*]+\s+[\d,]+\s+([\d,]+(?:\.\d+)?)\s+[\d,]+(?:\.\d+)?',
        full_text)
    if m:
        try:
            result['unit_price'] = float(m.group(1).replace(',', ''))
        except ValueError:
            pass

    # Grade / specification line — look for known grade patterns
    grade_pat = re.compile(
        r'\b(SPCC[-\w]*|SAE\d+[A-Z]?|Q\d+[A-Z]?|[A-Z]{2,}\d+[-\w]*)\b')
    grades_found = grade_pat.findall(full_text)
    if grades_found:
        result['detected_grade'] = grades_found[0]

    # Price terms — CFR / FOB / CIF line
    m = re.search(r'PRICE\s+TERMS\s*:\s*([^\n]+)', full_text, re.I)
    if m:
        result['price_terms'] = m.group(1).strip().rstrip('.')

    # Country of origin
    result['country_of_origin'] = 'CHINA'

    return result


# ── Excel parsing ────────────────────────────────────────────────────────────
def _format_size(raw: str) -> str:
    """'1.2*1219' → '1.20X1219',  '5.5' → '5.5'"""
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
    """Check whether a DataFrame looks like an ERP export."""
    cols = [str(c) for c in df.columns]
    return any('牌号' in c or 'grade' in c.lower() for c in cols) and \
           any('重量' in c or 'weight' in c.lower() for c in cols)


def _is_detail(df) -> bool:
    """Check whether a DataFrame looks like a 明细 sheet."""
    cols = ' '.join(str(c) for c in df.columns)
    return '净重' in cols or '毛重' in cols


def parse_excel_input(file_bytes: bytes, filename: str) -> dict:
    """
    Return:
      {
        'type': 'erp' | 'detail' | 'unknown',
        'grades': [
          {
            'grade':          str,
            'size':           str,
            'quantity_mt':    float,   # net weight
            'num_coils':      int,
            'gross_weight_mt': float,
            'unit_price':     float or None,
          }, ...
        ]
      }
    """
    try:
        import pandas as pd
    except ImportError:
        return {'type': 'unknown', 'grades': []}

    # Read file
    fname = filename.lower()
    try:
        if fname.endswith('.xlsx'):
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=0, engine='openpyxl')
        else:  # .xls
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')
    except Exception as e:
        return {'type': 'unknown', 'grades': [], 'error': str(e)}

    # ── ERP format detection ─────────────────────────────────────────────
    if fname.endswith('.xlsx') and _is_erp(df_raw):
        return _parse_erp(df_raw)

    # ── 明细 (.xls) format ───────────────────────────────────────────────
    return _parse_detail(df_raw)


def _parse_erp(df) -> dict:
    """
    ERP columns: 物流号, 物料名称, 车号, 出库库房, 合同号, 客户名称,
                 牌号, 钢板号, 简化后钢板号, 熔炉号, 质检批号, 简化后批号,
                 重量, 厚, 宽, 长, 出库时间
    Each row = 1 coil.
    """
    # Identify key columns by name
    cols = list(df.columns)
    grade_col  = next((c for c in cols if '牌号' in str(c)), None)
    weight_col = next((c for c in cols if '重量' in str(c)), None)
    thick_col  = next((c for c in cols if '厚'   in str(c)), None)

    if grade_col is None or weight_col is None:
        return {'type': 'unknown', 'grades': []}

    # Drop rows where grade is empty
    df = df.dropna(subset=[grade_col])
    df = df[df[grade_col].astype(str).str.strip() != '']

    grades_out = []
    for grade, grp in df.groupby(grade_col, sort=False):
        size_val = ''
        if thick_col and str(grp[thick_col].iloc[0]).strip() not in ('', '0', '0.0', 'nan'):
            size_val = _format_size(grp[thick_col].iloc[0])

        net_weight = round(float(grp[weight_col].sum()), 3)
        num_coils  = len(grp)

        grades_out.append({
            'grade':           str(grade).strip(),
            'size':            size_val,
            'quantity_mt':     net_weight,
            'num_coils':       num_coils,
            'gross_weight_mt': net_weight,  # wire rod: gross = net
            'unit_price':      None,
        })

    return {'type': 'erp', 'grades': grades_out}


def _parse_detail(df) -> dict:
    """
    明细 format (xls).  Example header (row 1):
      (blank), (blank), 卷号, 规格, 级别, 净重, 毛重
    Data rows start at row 2; last row may be a total row.
    """
    # Find header row (first row that contains '净重' or '卷号')
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v) for v in row]
        if any('净重' in v or '卷号' in v or '规格' in v for v in vals):
            header_row = i
            break

    if header_row is None:
        # Fallback: assume structure by position
        net_col   = 5
        gross_col = 6
        spec_col  = 3
        start_row = 2
        end_row   = len(df) - 1  # exclude last (total) row
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
        start_row = header_row + 1
        end_row   = len(df)  # will filter total rows below

    data_rows = df.iloc[start_row:end_row]

    # Drop total rows (detect by non-numeric net weight or 'coil' keyword)
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

    # Group by spec (there may be multiple specs in one 明细)
    def get_spec(row):
        if spec_col is not None:
            return str(row.iloc[spec_col]).strip()
        return ''

    groups: dict = {}
    for _, row in data_rows.iterrows():
        spec = get_spec(row)
        if spec not in groups:
            groups[spec] = {'coils': 0, 'net': 0.0, 'gross': 0.0}
        groups[spec]['coils'] += 1
        try:
            groups[spec]['net'] = round(groups[spec]['net'] + float(row.iloc[net_col]), 3)
        except (TypeError, ValueError):
            pass
        if gross_col is not None:
            try:
                groups[spec]['gross'] = round(
                    groups[spec]['gross'] + float(row.iloc[gross_col]), 3)
            except (TypeError, ValueError):
                pass

    grades_out = []
    import math
    for spec, g in groups.items():
        net   = round(g['net'],   3)
        if math.isnan(net) or net == 0:
            continue
        gross = round(g['gross'], 3) if g['gross'] else net
        grades_out.append({
            'grade':           '',              # from contract / user form
            'size':            _format_size(spec),
            'quantity_mt':     net,
            'num_coils':       g['coils'],
            'gross_weight_mt': gross,
            'unit_price':      None,
        })

    return {'type': 'detail', 'grades': grades_out}
