"""Strict CSV parsing and deterministic reconciliation using integer cents."""
import json
import re

MAX_TEXT_BYTES = 2 * 1024 * 1024


def parse(text):
    if len(text.encode('utf-8')) > MAX_TEXT_BYTES:
        raise ValueError('每份输入请控制在 2 MB 以内。')
    text = text.removeprefix('\ufeff')
    quoted = False
    comma = tab = 0
    index = 0
    while index < len(text):
        char = text[index]
        if char == '"':
            if quoted and index + 1 < len(text) and text[index + 1] == '"':
                index += 1
            else:
                quoted = not quoted
        elif not quoted:
            if char in '\n\r':
                break
            comma += char == ','
            tab += char == '\t'
        index += 1
    separator = '\t' if tab > comma else ','
    rows, row, field = [], [], []
    state = 0  # unquoted, quoted, closed quote

    def push_field():
        nonlocal field, state
        row.append(''.join(field))
        field, state = [], 0

    def push_row():
        nonlocal row
        push_field()
        if any(value.strip() for value in row):
            rows.append(row)
        row = []

    index = 0
    while index < len(text):
        char = text[index]
        if state == 1:
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    field.append('"')
                    index += 1
                else:
                    state = 2
            else:
                field.append(char)
        elif char == separator:
            push_field()
        elif char in '\n\r':
            if char == '\r' and index + 1 < len(text) and text[index + 1] == '\n':
                index += 1
            push_row()
        elif state == 2:
            if char != ' ' and not (char == '\t' and separator != '\t'):
                raise ValueError('引号字段结束后存在多余字符。')
        elif char == '"':
            if field:
                raise ValueError('引号应位于字段开头。')
            state = 1
        else:
            field.append(char)
        index += 1
    if state == 1:
        raise ValueError('CSV 引号未闭合。')
    if field or row or state == 2:
        push_row()
    if not rows:
        raise ValueError('请提供表头和明细记录。')
    headers = [value.strip() for value in rows.pop(0)]
    if any(not name for name in headers) or len(set(headers)) != len(headers):
        raise ValueError('表头不能为空或重复。')
    if len(headers) > 50 or len(rows) > 5000:
        raise ValueError('每份数据最多 50 列、5000 行。')
    for number, values in enumerate(rows, 2):
        if len(values) != len(headers):
            raise ValueError(f'第 {number} 条表格记录的列数与表头不一致。')
    if not rows:
        raise ValueError('表头下面还没有明细。')
    return {'headers': headers, 'rows': rows}


def money(value):
    text = str(value).strip()
    if not text or len(text) > 100:
        return None
    if re.fullmatch(r'\([\d,.]+\)', text, re.ASCII):
        text = '-' + text[1:-1]
    if not re.fullmatch(r'[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,2})?', text, re.ASCII):
        return None
    negative = text.startswith('-')
    parts = text.replace(',', '').lstrip('+-').split('.')
    cents = int(parts[0]) * 100 + int(parts[1].ljust(2, '0') if len(parts) == 2 else '0')
    return -cents if negative else cents


def fmt(cents, grouping=True):
    if cents is None:
        return '不可计算'
    whole, fraction = divmod(abs(cents), 100)
    whole_text = format(whole, ',d') if grouping else str(whole)
    return ('-' if cents < 0 else '') + whole_text + f'.{fraction:02d}'


def reconcile(left, right, left_keys, left_amount, right_keys, right_amount):
    def prepare(data, keys, amount):
        keys = keys if isinstance(keys, list) else [keys]
        valid_index = lambda value: type(value) is int and 0 <= value < len(data['headers'])
        if (not keys or len(set(keys)) != len(keys) or not valid_index(amount)
                or any(not valid_index(key) or key == amount for key in keys)):
            raise ValueError('编号列与金额列需要选择不同的有效列。')
        groups, invalid, total, bad = {}, [], 0, 0
        for index, values in enumerate(data['rows']):
            parts = [values[key].strip() for key in keys]
            identity = '' if any(not value for value in parts) else (parts[0] if len(parts) == 1 else json.dumps(parts, ensure_ascii=False, separators=(',', ':')))
            cents = money(values[amount])
            record = {'line': data.get('sourceLines', [])[index] if 'sourceLines' in data else index + 2,
                      'id': identity, 'raw': values[amount], 'amount': cents, 'values': values}
            if cents is None:
                bad += 1
            else:
                total += cents
            if identity:
                groups.setdefault(identity, []).append(record)
            else:
                invalid.append(record)
        return groups, invalid, None if bad else total, bad

    lg, li, lt, lb = prepare(left, left_keys, left_amount)
    rg, ri, rt, rb = prepare(right, right_keys, right_amount)
    result = []
    for identity in dict.fromkeys([*lg, *rg]):
        a, b = lg.get(identity, []), rg.get(identity, [])
        delta = None
        if len(a) > 1 or len(b) > 1:
            status = '重复编号'
        elif any(item['amount'] is None for item in a + b):
            status = '无效金额'
        elif not a:
            status = '仅右表'
        elif not b:
            status = '仅左表'
        else:
            delta = a[0]['amount'] - b[0]['amount']
            status = '一致' if delta == 0 else '金额差异'
        result.append({'id': identity, 'status': status, 'left': a, 'right': b, 'delta': delta})
    for side, records in [('left', li), ('right', ri)]:
        for record in records:
            result.append({'id': '（空编号）', 'status': '空编号及无效金额' if record['amount'] is None else '空编号',
                           'left': [record] if side == 'left' else [], 'right': [record] if side == 'right' else [], 'delta': None})
    return {'result': result, 'leftTotal': lt, 'rightTotal': rt, 'leftBad': lb, 'rightBad': rb,
            'leftRows': len(left['rows']), 'rightRows': len(right['rows'])}


def export_csv(records):
    def cell(value):
        text = str(value)
        if re.match(r'\s*[=+\-@]', text):
            text = "'" + text
        return '"' + text.replace('"', '""') + '"'
    rows = [['编号', '状态', '左表记录序号', '右表记录序号', '左表原金额', '右表原金额', '左减右差额']]
    for record in records:
        rows.append([record['id'], record['status'],
                     ' / '.join(str(item['line']) for item in record['left']),
                     ' / '.join(str(item['line']) for item in record['right']),
                     ' / '.join(item['raw'] for item in record['left']),
                     ' / '.join(item['raw'] for item in record['right']),
                     '待核实' if record['delta'] is None else fmt(record['delta'])])
    return '\ufeff' + '\r\n'.join(','.join(cell(value) for value in row) for row in rows)
