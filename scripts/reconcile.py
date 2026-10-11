#!/usr/bin/env python3
"""Local CSV/XLSX reconciliation CLI; Python is the only runtime."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile

import core

VERSION = '0.8.5'


def fingerprint(path, limit=None):
    digest = hashlib.sha256()
    count = 0
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(65536), b''):
            count += len(block)
            if limit is not None and count > limit:
                raise ValueError('输入超过大小限制')
            digest.update(block)
    return digest.hexdigest()


def load_table(file, sheet=None, header_row='1'):
    path = Path(file)
    if not str(header_row).isascii() or not str(header_row).isdigit() or int(header_row) < 1:
        raise ValueError('表头行必须是正整数')
    if path.stat().st_size > 10 * 1024 * 1024:
        raise ValueError('输入超过10MB')
    before = fingerprint(path, 10 * 1024 * 1024)
    metadata = None
    if path.suffix.lower() == '.xlsx':
        result = subprocess.run([sys.executable, '-X', 'utf8', str(Path(__file__).with_name('xlsx_read.py')),
                                 str(path), sheet or '', str(header_row)],
                                capture_output=True, text=True, encoding='utf-8', timeout=30)
        if result.returncode:
            raise ValueError(result.stderr.strip())
        metadata = json.loads(result.stdout)
        text = metadata['csv']
    elif path.suffix.lower() in {'.csv', '.tsv'}:
        if int(header_row) != 1 or sheet:
            raise ValueError('工作表和自定义表头行仅支持XLSX')
        if path.stat().st_size > core.MAX_TEXT_BYTES:
            raise ValueError('CSV超过2MB')
        text = path.read_bytes().decode('utf-8')
    else:
        raise ValueError('支持UTF-8 CSV、TSV、XLSX；旧xls请先另存')
    table = core.parse(text)
    if metadata:
        if len(metadata['sourceLines']) != len(table['rows']):
            raise ValueError('XLSX来源行数量不一致')
        table['sourceLines'] = metadata['sourceLines']
    if before != fingerprint(path, 10 * 1024 * 1024):
        raise ValueError('读取过程中输入文件发生变化')
    return table, {'file': path.name, 'sha256': before, 'sheet': metadata['sheet'] if metadata else None,
                   'headerRow': int(header_row), 'rows': len(table['rows'])}


def serializable_report(report):
    def amount(value):
        return None if value is None else str(value)
    output = dict(report)
    for key in ['leftTotal', 'rightTotal']:
        output[key] = amount(report[key])
    output['result'] = []
    for record in report['result']:
        item = {**record, 'delta': amount(record['delta'])}
        for side in ['left', 'right']:
            item[side] = [{**source, 'amount': amount(source['amount'])} for source in record[side]]
        output['result'].append(item)
    return output


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    parser = argparse.ArgumentParser(description='财务明细核对 '+VERSION)
    for name in ['left', 'right', 'left-key', 'right-key', 'left-keys', 'right-keys',
                 'left-amount', 'right-amount', 'left-sheet', 'right-sheet', 'out']:
        parser.add_argument('--'+name)
    for side in ['left', 'right']:
        parser.add_argument('--'+side+'-header-row', default='1')
    parser.add_argument('--inspect', action='store_true')
    parser.add_argument('--single-key-confirmed', action='store_true')
    if not argv:
        parser.print_help()
        return
    seen = set()
    for token in argv:
        if token.startswith('--'):
            option = token.split('=', 1)[0]
            if option in seen:
                raise ValueError('重复参数：'+option)
            seen.add(option)
    args = vars(parser.parse_args(argv))
    tables, inputs = {}, {}
    for side in ['left', 'right']:
        if not args[side]:
            raise ValueError('缺少 --'+side)
        tables[side], inputs[side] = load_table(args[side], args[side+'_sheet'], args[side+'_header_row'])
    if args['inspect']:
        print(json.dumps({side: {'headers': table['headers'], 'rows': len(table['rows'])}
                          for side, table in tables.items()}, ensure_ascii=False, indent=2))
        return

    def keys(side):
        multiple, single = args[side+'_keys'], args[side+'_key']
        if multiple and single:
            raise ValueError('单列与联合键参数二选一')
        names = json.loads(multiple) if multiple else [single]
        if not isinstance(names, list) or not names or any(not isinstance(name, str) or name not in tables[side]['headers'] for name in names):
            raise ValueError('编号字段必须是现有列名组成的JSON数组')
        return [tables[side]['headers'].index(name) for name in names]

    lk, rk = keys('left'), keys('right')
    if len(lk) != len(rk):
        raise ValueError('两侧联合键字段数量及顺序必须对应')
    scope_names = {'客户', '客户名称', '公司', '公司名称', '客户编码', '公司编码'}
    if len(lk) == 1 and not args['single_key_confirmed'] and all(scope_names.intersection(tables[side]['headers']) for side in tables):
        raise ValueError('两表含客户或公司字段，请联合匹配；已确认编号全局唯一时才加 --single-key-confirmed')
    amount_columns = []
    for side in ['left', 'right']:
        if args[side+'_amount'] not in tables[side]['headers']:
            raise ValueError('金额列不存在：'+side)
        amount_columns.append(tables[side]['headers'].index(args[side+'_amount']))
    report = core.reconcile(tables['left'], tables['right'], lk, amount_columns[0], rk, amount_columns[1])
    coverage = {}
    for side in tables:
        records = [item for group in report['result'] for item in group[side]]
        expected = len(tables[side]['rows'])
        if len(records) != expected or len({item['line'] for item in records}) != expected:
            raise ValueError('逐条覆盖检查失败：'+side)
        if fingerprint(args[side], 10 * 1024 * 1024) != inputs[side]['sha256']:
            raise ValueError('生成报告前输入发生变化：'+side)
        coverage[side] = {'inputRows': expected, 'reportedRows': len(records), 'everyRecordExactlyOnce': True}
    if not args['out']:
        raise ValueError('缺少 --out（新目录）')
    out = Path(args['out'])
    if out.exists():
        raise ValueError('结果目录已存在，请使用新目录')
    counts = {}
    for record in report['result']:
        counts[record['status']] = counts.get(record['status'], 0) + 1
    summary = {'leftRows': report['leftRows'], 'rightRows': report['rightRows'],
               'leftTotal': core.fmt(report['leftTotal']), 'rightTotal': core.fmt(report['rightTotal']),
               'counts': counts, 'moneyUnit': '输入金额单位，保留两位小数；JSON 整数金额单位为分',
               'differenceRule': '左表减右表；重复或异常不计算差额'}
    with tempfile.TemporaryDirectory(dir=out.parent) as staging:
        stage = Path(staging)/'result'
        stage.mkdir()
        for name, content in [('差异表.csv', core.export_csv([r for r in report['result'] if r['status'] != '一致'])),
                              ('完整核对表.csv', core.export_csv(report['result'])),
                              ('核对记录.json', json.dumps(serializable_report(report), ensure_ascii=False, indent=2)),
                              ('摘要.json', json.dumps(summary, ensure_ascii=False, indent=2))]:
            with (stage/name).open('w', encoding='utf-8', newline='') as stream:
                stream.write(content)
        outputs = [{'file': name, 'sha256': fingerprint(stage/name)}
                   for name in ['差异表.csv', '完整核对表.csv', '核对记录.json', '摘要.json']]
        manifest = {'skill': 'finance-reconcile', 'version': VERSION,
                    'runAt': datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'), 'inputs': inputs,
                    'fieldMapping': {'left': {'keys': [tables['left']['headers'][i] for i in lk], 'amount': args['left_amount']},
                                     'right': {'keys': [tables['right']['headers'][i] for i in rk], 'amount': args['right_amount']}},
                    'rules': {'id': 'ordered-tuple-trim-case-sensitive', 'amount': 'integer-cents-exact',
                              'duplicates': 'retain-all-no-aggregation', 'delta': 'left-minus-right'},
                    'coverage': coverage, 'outputs': outputs}
        (stage/'运行清单.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
        stage.rename(out)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print('核对失败：'+str(error), file=sys.stderr)
        sys.exit(1)
