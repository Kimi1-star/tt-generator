"""
TT (Commercial Invoice / Packing List / Shipment Advice) Excel generator.
"""
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ── Fixed company / bank constants ─────────────────────────────────────────
COMPANY_NAME = "SHARPMAX INTERNATIONAL (HONG KONG) CO., LIMITED"
COMPANY_ADDR = ("ROOM 704, 7/F., TOWER A, NEW MANDARIN PLAZA, "
                "14 SCIENCE MUSEUM ROAD, TST EAST, KOWLOON, HONG KONG")
BENEFICIARY  = "SHARPMAX INTERNATIONAL (HONGKONG) CO.,LIMITED"
BANK_NAME    = "CHINA ZHESHANG BANK"
BANK_ADDR    = "NO.288 QINGCHUN ROAD HANG ZHOU,ZHEJIANG,CHINA"
SWIFT_CODE   = "ZJCBCN2N"
ACCOUNT_NO   = "NRA3310010711420100000726"

# ── Number → English words ──────────────────────────────────────────────────
_ONES = ['', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT',
         'NINE', 'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN',
         'SIXTEEN', 'SEVENTEEN', 'EIGHTEEN', 'NINETEEN']
_TENS = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY',
         'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY']

def _below_thousand(n: int) -> str:
    if n == 0:
        return ''
    if n < 20:
        return _ONES[n]
    if n < 100:
        return _TENS[n // 10] + (' ' + _ONES[n % 10] if n % 10 else '')
    h   = _ONES[n // 100] + ' HUNDRED'
    rem = n % 100
    return h + (' AND ' + _below_thousand(rem) if rem else '')

def amount_to_words(amount: float) -> str:
    dollars = int(amount)
    cents   = round((amount - dollars) * 100)

    def _full(n):
        if n == 0:
            return 'ZERO'
        parts = []
        if n >= 1_000_000:
            parts.append(_below_thousand(n // 1_000_000) + ' MILLION')
            n %= 1_000_000
        if n >= 1_000:
            parts.append(_below_thousand(n // 1_000) + ' THOUSAND')
            n %= 1_000
        if n > 0:
            parts.append(_below_thousand(n))
        return ' '.join(parts)

    txt = 'SAY US DOLLARS ' + _full(dollars)
    if cents:
        txt += ' AND CENTS ' + _below_thousand(cents)
    return txt + ' ONLY'


# ── Style helpers ───────────────────────────────────────────────────────────
def _thin_border():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

def _font(bold=False, size=10):
    return Font(name='Arial', bold=bold, size=size)

def _align(h='left', wrap=True):
    return Alignment(horizontal=h, vertical='center', wrap_text=wrap)

def _w(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width

def _h(ws, row, height):
    ws.row_dimensions[row].height = height

def _cell(ws, row, col, value='', bold=False, size=10,
          h_align='left', wrap=True, border=False, fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = _font(bold, size)
    c.alignment = _align(h_align, wrap)
    if border:
        c.border = _thin_border()
    if fmt:
        c.number_format = fmt
    return c

def _merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1,
                   end_row=r2, end_column=c2)


# ── CI Sheet ────────────────────────────────────────────────────────────────
def _build_ci(ws, d):
    # Column widths (6 columns A–F)
    for col, w in zip('ABCDEF', [46, 12, 12, 12, 12, 16]):
        _w(ws, col, w)

    r = 1
    # ── Company header
    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, COMPANY_NAME, bold=True, size=11, h_align='center')
    _h(ws, r, 18); r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, COMPANY_ADDR, size=9, h_align='center')
    _h(ws, r, 15); r += 1

    r += 1  # blank

    # ── Title
    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, 'COMMERCIAL INVOICE', bold=True, size=14, h_align='center')
    _h(ws, r, 22); r += 1

    r += 2  # blank rows

    # ── ORIGINAL
    _cell(ws, r, 6, 'ORIGINAL', bold=True, h_align='right')
    r += 1

    # ── TO:
    _cell(ws, r, 1, 'TO:', bold=True)
    r += 1

    # ── Buyer info | DATE
    buyer_lines = [d.get('buyer_name', '')]
    if d.get('buyer_address'):
        buyer_lines.append('ADD: ' + d['buyer_address'])
    if d.get('buyer_ref'):
        buyer_lines.append(d['buyer_ref'])
    _cell(ws, r, 1, '\n'.join(buyer_lines))
    _cell(ws, r, 6, f"DATE: {d['invoice_date']}", h_align='right')
    _h(ws, r, 50); r += 1

    # ── INVOICE NO.
    _cell(ws, r, 6, f"INVOICE NO.: {d['invoice_no']}", h_align='right')
    r += 1

    r += 3  # blank rows

    # ── Description section
    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, 'DESCRIPTION OF GOODS & SERVICES', bold=True)
    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, f"COMMODITY: {d['commodity']}")
    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, f"SALES CONTRACT NO.: {d['contract_no']}")
    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, f"COUNTRY OF ORIGIN: {d.get('country_of_origin', 'CHINA')}")
    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, f"PRICE TERMS: {d['price_terms']}")
    r += 1

    r += 1  # blank

    # ── Table header (2 rows)
    hdrs1 = ['', 'GRADE', 'SIZE', 'QUANTITY', 'UNIT PRICE', 'AMOUNT']
    hdrs2 = ['', '',      '(MM)', '(MT)',     '(USD/MT)',   '(USD)']
    for col, (h1, h2) in enumerate(zip(hdrs1, hdrs2), 1):
        _cell(ws, r,   col, h1, bold=True, h_align='center', border=True)
        _cell(ws, r+1, col, h2, bold=True, h_align='center', border=True)
    r += 2

    # ── Data rows
    total_qty = 0.0
    total_amt = 0.0
    for g in d['grades']:
        qty   = round(float(g['quantity_mt']), 3)
        price = round(float(g['unit_price']),  2)
        amt   = round(qty * price, 2)
        total_qty = round(total_qty + qty, 3)
        total_amt = round(total_amt + amt, 2)

        row_vals = ['', g['grade'], str(g['size']), qty, price, amt]
        fmts     = [None, None, None, '#,##0.000', '#,##0.00', '#,##0.00']
        for col, (v, f) in enumerate(zip(row_vals, fmts), 1):
            align = 'right' if col >= 4 else 'center'
            _cell(ws, r, col, v, h_align=align, border=True, fmt=f)
        r += 1

    # ── TOTAL
    for col, v in enumerate(['TOTAL', '', '', total_qty, '', total_amt], 1):
        f = ('#,##0.000' if col == 4 else '#,##0.00' if col == 6 else None)
        a = ('right' if col in (4, 6) else 'center')
        _cell(ws, r, col, v, bold=True, h_align=a, border=True, fmt=f)
    r += 1

    # ── Prepayment / balance (has_prepayment path)
    if d.get('has_prepayment') and d.get('prepayment'):
        prepayment = round(float(d['prepayment']), 2)
        balance    = round(total_amt - prepayment, 2)
        say_amount = balance

        _cell(ws, r, 1, 'DEDUCTION (DOWN PAYMENT)', bold=True)
        _cell(ws, r, 6, prepayment, bold=False, h_align='right', fmt='#,##0.00')
        r += 1

        _cell(ws, r, 1, 'BALANCE', bold=True)
        _cell(ws, r, 6, balance, bold=True, h_align='right', fmt='#,##0.00')
        r += 1
    else:
        say_amount = total_amt

    # ── SAY row
    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, amount_to_words(say_amount))
    _h(ws, r, 28); r += 1

    # ── Path-specific tail
    if d.get('has_prepayment') and d.get('prepayment'):
        r += 1  # blank
        # BENEFICIARY first, then FOB/FREIGHT
        _merge(ws, r, 1, r, 6)
        _cell(ws, r, 1, f'BENEFICIARY: {BENEFICIARY}')
        r += 1
        _merge(ws, r, 1, r, 6)
        _cell(ws, r, 1, f'ADD: {COMPANY_ADDR}')
        r += 1
        for label, val in [('BANK NAME', BANK_NAME), ('BANK ADDRESS', BANK_ADDR),
                            ('SWIFT CODE', SWIFT_CODE), ('ACCOUNT NO.', ACCOUNT_NO)]:
            _merge(ws, r, 1, r, 6)
            _cell(ws, r, 1, f'{label}:{val}' if label == 'SWIFT CODE'
                  else f'{label}: {val}')
            r += 1

        fob = round(total_amt - float(d.get('freight') or 0), 2)
        _cell(ws, r, 1, 'FOB VALUE: USD')
        _cell(ws, r, 3, fob,                     h_align='right', fmt='#,##0.00')
        r += 1
        _cell(ws, r, 1, 'FREIGHT AMOUNT: USD')
        _cell(ws, r, 3, float(d.get('freight') or 0), h_align='right', fmt='#,##0.00')
        r += 1

    else:
        # Payment clause, then FOB/FREIGHT, then BENEFICIARY
        pay_txt = ("THE INVOICE AMOUNT SHOULD BE PAID WITHIN 5 WORKING DAYS AGAINST "
                   "BL SCANNED COPY, COMMERCIAL INVOICE, PACKING LIST AT THE FIRST "
                   "VERSION SENT BY THE SELLER THROUGH THE EMAIL.")
        _merge(ws, r, 1, r, 6)
        _cell(ws, r, 1, pay_txt)
        _h(ws, r, 35); r += 1

        fob = round(total_amt - float(d.get('freight') or 0), 2)
        _cell(ws, r, 1, 'FOB VALUE: USD')
        _cell(ws, r, 2, fob,                     h_align='right', fmt='#,##0.00')
        r += 1
        _cell(ws, r, 1, 'FREIGHT AMOUNT: USD')
        _cell(ws, r, 2, float(d.get('freight') or 0), h_align='right', fmt='#,##0.00')
        r += 2  # blank

        _merge(ws, r, 1, r, 6)
        _cell(ws, r, 1, f'BENEFICIARY: {BENEFICIARY}')
        r += 1
        _merge(ws, r, 1, r, 6)
        _cell(ws, r, 1, f'ADD: {COMPANY_ADDR}')
        r += 1
        for label, val in [('BANK NAME', BANK_NAME), ('BANK ADDRESS', BANK_ADDR),
                            ('SWIFT CODE', SWIFT_CODE), ('ACCOUNT NO.', ACCOUNT_NO)]:
            _merge(ws, r, 1, r, 6)
            _cell(ws, r, 1, f'{label}:{val}' if label == 'SWIFT CODE'
                  else f'{label}: {val}')
            r += 1

    r += 2  # blank rows
    _cell(ws, r, 4, COMPANY_NAME, bold=True, h_align='center')


# ── PL Sheet ────────────────────────────────────────────────────────────────
def _build_pl(ws, d):
    for col, w in zip('ABCDEF', [28, 14, 14, 14, 16, 16]):
        _w(ws, col, w)

    r = 1
    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, COMPANY_NAME, bold=True, size=11, h_align='center')
    _h(ws, r, 18); r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, COMPANY_ADDR, size=9, h_align='center')
    _h(ws, r, 15); r += 1

    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, 'PACKING LIST', bold=True, size=14, h_align='center')
    _h(ws, r, 22); r += 1

    r += 2

    _cell(ws, r, 6, 'ORIGINAL', bold=True, h_align='right'); r += 1
    _cell(ws, r, 1, 'TO:', bold=True); r += 1

    buyer_lines = [d.get('buyer_name', '')]
    if d.get('buyer_address'):
        buyer_lines.append('ADD: ' + d['buyer_address'])
    if d.get('buyer_ref'):
        buyer_lines.append(d['buyer_ref'])
    _cell(ws, r, 1, '\n'.join(buyer_lines))
    _cell(ws, r, 6, f"DATE: {d['invoice_date']}", h_align='right')
    _h(ws, r, 50); r += 1

    _cell(ws, r, 6, f"INVOICE NO.: {d['invoice_no']}", h_align='right')
    r += 1

    r += 3

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, 'DESCRIPTION OF GOODS & SERVICES', bold=True)
    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, f"COMMODITY: {d['commodity']}")
    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, f"SALES CONTRACT NO.: {d['contract_no']}")
    r += 1

    _merge(ws, r, 1, r, 6)
    _cell(ws, r, 1, f"COUNTRY OF ORIGIN: {d.get('country_of_origin', 'CHINA')}")
    r += 1

    r += 1

    # Table header (2 rows)
    hdrs1 = ['GRADE', 'SIZE',  'NUMBER OF', 'NET WEIGHT /\nQUANTITY', 'GROSS WEIGHT', '']
    hdrs2 = ['',      '(MM)',  'COILS',     '(MT)',                   '(MT)',          '']
    for col, (h1, h2) in enumerate(zip(hdrs1, hdrs2), 1):
        _cell(ws, r,   col, h1, bold=True, h_align='center', border=True)
        _cell(ws, r+1, col, h2, bold=True, h_align='center', border=True)
    r += 2

    total_coils = 0
    total_net   = 0.0
    total_gross = 0.0

    for g in d['grades']:
        coils = int(g.get('num_coils', 0))
        net   = round(float(g['quantity_mt']), 3)
        gross = round(float(g.get('gross_weight_mt') or net), 3)
        total_coils += coils
        total_net    = round(total_net  + net,   3)
        total_gross  = round(total_gross + gross, 3)

        row_vals = [g['grade'], str(g['size']), coils, net, gross, '']
        fmts     = [None, None, '#,##0', '#,##0.000', '#,##0.000', None]
        for col, (v, f) in enumerate(zip(row_vals, fmts), 1):
            a = 'right' if col >= 3 else 'center'
            _cell(ws, r, col, v, h_align=a, border=True, fmt=f)
        r += 1

    for col, (v, f) in enumerate(
            [('TOTAL', None), ('', None), (total_coils, '#,##0'),
             (total_net, '#,##0.000'), (total_gross, '#,##0.000'), ('', None)], 1):
        a = 'right' if col >= 3 else 'center'
        _cell(ws, r, col, v, bold=True, h_align=a, border=True, fmt=f)
    r += 1

    r += 3
    _cell(ws, r, 3, COMPANY_NAME, bold=True, h_align='center')


# ── SA Sheet ────────────────────────────────────────────────────────────────
def _build_sa(ws, d):
    for col, w in zip('ABCDEFG', [30, 14, 14, 14, 14, 14, 14]):
        _w(ws, col, w)

    r = 1
    _merge(ws, r, 1, r, 7)
    _cell(ws, r, 1, COMPANY_NAME, bold=True, size=11, h_align='center')
    _h(ws, r, 18); r += 1

    _merge(ws, r, 1, r, 7)
    _cell(ws, r, 1, COMPANY_ADDR, size=9, h_align='center')
    _h(ws, r, 15); r += 1

    r += 1

    _merge(ws, r, 1, r, 7)
    _cell(ws, r, 1, 'SHIPMENT ADVICE', bold=True, size=14, h_align='center')
    _h(ws, r, 22); r += 1

    r += 2

    _cell(ws, r, 7, 'ORIGINAL', bold=True, h_align='right'); r += 1
    _cell(ws, r, 1, 'TO:', bold=True); r += 1

    buyer_lines = [d.get('buyer_name', '')]
    if d.get('buyer_address'):
        buyer_lines.append('ADD: ' + d['buyer_address'])
    if d.get('buyer_ref'):
        buyer_lines.append(d['buyer_ref'])
    _cell(ws, r, 1, '\n'.join(buyer_lines))
    _cell(ws, r, 7, f"DATE: {d['invoice_date']}", h_align='right')
    _h(ws, r, 50); r += 1

    _cell(ws, r, 7, f"INVOICE NO.: {d['invoice_no']}", h_align='right')
    r += 1

    r += 3

    _merge(ws, r, 1, r, 7)
    _cell(ws, r, 1, 'SHIPMENT DETAILS', bold=True)
    r += 1

    _merge(ws, r, 1, r, 7)
    _cell(ws, r, 1, f"COMMODITY: {d['commodity']}")
    r += 1

    total_qty = round(sum(float(g['quantity_mt']) for g in d['grades']), 3)
    total_amt = round(sum(float(g['quantity_mt']) * float(g['unit_price'])
                         for g in d['grades']), 2)

    _cell(ws, r, 1, 'QUANTITY LOADED: MT')
    _cell(ws, r, 2, total_qty, h_align='right', fmt='#,##0.000')
    r += 1

    _cell(ws, r, 1, 'INVOICE VALUE: USD')
    _cell(ws, r, 2, total_amt, h_align='right', fmt='#,##0.00')
    r += 1

    _cell(ws, r, 1, f"NAME OF VESSEL: {d.get('vessel', '')}")
    r += 1

    _cell(ws, r, 1, f"PORT OF LOADING: {d.get('loading_port', '')}")
    r += 1

    eta_label = d.get('eta_etd_label', 'ETA').upper()
    _cell(ws, r, 1, f"{eta_label}: {d.get('eta_etd_date', '')}")
    r += 1

    r += 15  # spacing before signature
    _cell(ws, r, 3, COMPANY_NAME, bold=True, h_align='center')


# ── Entry point ─────────────────────────────────────────────────────────────
def generate_tt_excel(data: dict) -> bytes:
    """
    Build an Excel workbook (CI / PL / SA) and return bytes.
    """
    wb = Workbook()
    wb.remove(wb.active)

    ci_ws = wb.create_sheet('CI')
    _build_ci(ci_ws, data)

    pl_ws = wb.create_sheet('PL')
    _build_pl(pl_ws, data)

    sa_ws = wb.create_sheet('SA')
    _build_sa(sa_ws, data)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
