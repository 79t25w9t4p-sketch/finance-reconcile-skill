#!/usr/bin/env python3
"""Repeatable local reconciliation and review workbook workflow."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile
from decimal import Decimal

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
VERSION = '0.6.0'
HEADERS = ['编号 / 联合键', '核对结果', '本次变化', '处理状态', '处理意见', '左表来源行', '右表来源行', '左表原金额', '右表原金额', '左减右差额', '记录标识', '内容指纹']
STATES = {'待处理', '处理中', '已解决', '需重新复核', '无需处理'}


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def read_previous(path, profile_hash):
    if not path:
        return {}, []
    path = Path(path)
    if path.stat().st_size > 10 * 1024 * 1024:
        raise ValueError('上次工作簿超过10MB')
    with zipfile.ZipFile(path) as z:
        if sum(x.file_size for x in z.infolist()) > 40 * 1024 * 1024:
            raise ValueError('上次工作簿解压后超过40MB')
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        if wb['_meta'].max_row != 2 or wb['_meta'].max_column != 2:
            raise ValueError('工作簿标识结构无效')
        meta = dict(wb['_meta'].iter_rows(values_only=True))
        if meta.get('format') != 'finance-review-v1' or meta.get('profile') != profile_hash:
            raise ValueError('上次工作簿的配置不一致；不能自动沿用意见')
        previous = {}
        history = []
        for name in ['待处理', '已匹配', '历史记录']:
            ws = wb[name]
            if ws.max_row > 10001 or ws.max_column > 12:
                raise ValueError('上次工作簿结构或记录数量异常')
            if list(next(ws.values)) != HEADERS:
                raise ValueError('请保留复核工作表的原表头')
            for cells in ws.iter_rows(min_row=2):
                if any(c.data_type == 'f' for c in cells):
                    raise ValueError('处理意见中请填写文本，不能使用公式')
                row = [c.value if c.value is not None else '' for c in cells]
                if not any(row):
                    continue
                if row[3] not in STATES:
                    raise ValueError('处理状态无效，请使用下拉选项')
                if name == '历史记录':
                    history.append(row)
                else:
                    if not row[10] or row[10] in previous:
                        raise ValueError('上次工作簿存在缺失或重复记录标识')
                    previous[row[10]] = row
        return previous, history
    finally:
        wb.close()


def text_cell(cell, value):
    # Treat all strings as literal text, including =, +, -, and @ prefixes.
    if isinstance(value, str):
        if len(value) > 32767:
            raise ValueError('单元格内容过长，请按较小范围核对')
        cell.value = value
        cell.data_type = 's'
    else:
        cell.value = value


def make_sheet(wb, name, rows):
    ws = wb.create_sheet(name)
    for row_number, row in enumerate([HEADERS] + rows, 1):
        for col, value in enumerate(row, 1):
            text_cell(ws.cell(row_number, col), value)
    ws.freeze_panes = 'F2'
    ws.auto_filter.ref = ws.dimensions
    widths = [34, 18, 20, 18, 48, 18, 18, 22, 22, 20, 20, 20]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.font = Font(color='FFFFFF', bold=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        for cell in row[3:5]:
            cell.fill = PatternFill('solid', fgColor='FFF2CC')
    ws.column_dimensions['K'].hidden = True
    ws.column_dimensions['L'].hidden = True
    if rows:
        dv = DataValidation(type='list', formula1='"待处理,处理中,已解决,需重新复核,无需处理"')
        dv.showErrorMessage = True
        dv.error = '请从下拉选项中选择处理状态'
        ws.add_data_validation(dv)
        dv.add(f'D2:D{len(rows)+1}')
    return ws


def build_workbook(report, summary, profile, previous_path, dest):
    profile_hash = digest(profile)
    previous, history = read_previous(previous_path, profile_hash)
    rows = []
    seen = set()
    changes = {}
    for index, item in enumerate(report['result']):
        ambiguous = '空编号' in item['status'] or item['status'] == '重复编号'
        identity = digest({'id': item['id'], 'blankIndex': index if '空编号' in item['status'] else None})
        contents = {'status': item['status'], 'delta': item['delta']}
        for side in ['left', 'right']:
            contents[side] = sorted([x.get('values', [x['id'], x['raw']]) for x in item[side]])
        fingerprint = digest(contents)
        old = previous.get(identity) if '空编号' not in item['status'] else None
        if '空编号' not in item['status']:
            seen.add(identity)
        note = old[4] if old else ''
        if old and old[11] == fingerprint and not ambiguous:
            change, state = '未变化', old[3]
        elif old:
            change, state = '需重新复核', '需重新复核'
        else:
            change = '新增' if previous_path else '首次核对'
            state = '无需处理' if item['status'] == '一致' else '待处理'
        changes[change] = changes.get(change, 0) + 1
        row = [item['id'], item['status'], change, state, note,
               ' / '.join(str(x['line']) for x in item['left']),
               ' / '.join(str(x['line']) for x in item['right']),
               ' / '.join(x['raw'] for x in item['left']),
               ' / '.join(x['raw'] for x in item['right']),
               '待核实' if item['delta'] is None else format(Decimal(item['delta']) / Decimal(100), '.2f'),
               identity, fingerprint]
        rows.append(row)
    for identity, old in previous.items():
        if identity not in seen:
            old[2] = '本次未出现'
            history.append(old)
    if len(history) > 10000:
        raise ValueError('历史记录超过10000条，请另存归档后开始新周期')
    priority = {'需重新复核': 0, '待处理': 1, '处理中': 2, '已解决': 3, '无需处理': 4}
    rows.sort(key=lambda r: (priority[r[3]], 0 if r[2] == '新增' else 1, r[0]))
    attention = sum(row[3] in {'需重新复核', '待处理', '处理中'} for row in rows)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '汇总'
    overview = [['基础对账 · 先看需要处理的事项', VERSION],
                ['本次待办', attention], ['需重新复核', sum(row[3] == '需重新复核' for row in rows)],
                ['新增待办', sum(row[2] == '新增' and row[3] != '无需处理' for row in rows)],
                ['意见沿用', sum(row[2] == '未变化' and bool(row[4]) for row in rows)],
                ['下一步', '打开待处理页；优先看需重新复核和新增条目，在黄色列填写状态和意见。'],
                ['统计口径', '生成时统计；手动修改处理状态后，汇总数字不实时更新，下次重跑重新统计。'],
                ['左表记录数', summary['leftRows']],
                ['右表记录数', summary['rightRows']], ['左表金额合计', summary['leftTotal']],
                ['右表金额合计', summary['rightTotal']], ['使用方式', '填写黄色列；已解决记录仍保留，可用处理状态筛选。下次把填写过的工作簿交给AI。'],
                ['重跑规则', '内容未变化才沿用状态；变化时保留意见并要求重新复核。'],
                ['历史记录', '本次未出现仅代表不在本次输入中，不代表问题已解决。']]
    overview += [['金额差异怎么查', '对照左右金额和来源行回查单据，不自动断言差额原因。'],
                 ['单边记录怎么查', '核实遗漏、期间和编号口径。'],
                 ['重复 / 无效值怎么查', '先确认明细粒度与原始数据，再重新核对。']]
    overview += list(summary['counts'].items()) + list(changes.items())
    for row in overview:
        ws.append(row)
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 85
    for row in ws:
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical='top')
    ws['A1'].font = Font(size=16, bold=True, color='17365D')
    ws.column_dimensions['A'].width = 38
    ws.freeze_panes = 'B2'
    ws.row_dimensions[1].height = 30
    for number in range(2, 6):
        ws.cell(number, 2).font = Font(size=16, bold=True, color='17365D')
        ws.cell(number, 2).fill = PatternFill('solid', fgColor='E8F1FA')
    for number in range(6, ws.max_row + 1):
        ws.row_dimensions[number].height = 32
    make_sheet(wb, '待处理', [r for r in rows if r[1] != '一致' or r[3] == '需重新复核'])
    make_sheet(wb, '已匹配', [r for r in rows if r[1] == '一致' and r[3] != '需重新复核'])
    make_sheet(wb, '历史记录', history)
    meta = wb.create_sheet('_meta')
    meta.append(['format', 'finance-review-v1'])
    meta.append(['profile', profile_hash])
    meta.sheet_state = 'hidden'
    wb.save(dest)
    return changes


def main():
    parser = argparse.ArgumentParser(description='保存核对配置，换表重跑，输出可填写复核意见的Excel')
    for name in ['config', 'save-config', 'left', 'right', 'out', 'previous', 'node', 'left-keys', 'right-keys', 'left-amount', 'right-amount', 'left-sheet', 'right-sheet', 'left-header-row', 'right-header-row']:
        parser.add_argument('--'+name)
    parser.add_argument('--single-key-confirmed', action='store_true')
    args = vars(parser.parse_args())
    for key in ['left', 'right', 'out']:
        if not args[key]:
            raise ValueError('缺少 --'+key)
    if bool(args['config']) == bool(args['save_config']):
        raise ValueError('首次用 --save-config；以后用 --config，二选一')
    fields = ['left_keys', 'right_keys', 'left_amount', 'right_amount', 'left_sheet', 'right_sheet', 'left_header_row', 'right_header_row', 'single_key_confirmed']
    if args['config']:
        if any(args[k] for k in fields):
            raise ValueError('复用配置时不能同时改映射；格式变化请创建新配置')
        profile = read_json(args['config'])
        if set(profile) != {'version', 'mapping', 'headers'} or profile['version'] != 1 or not set(profile['mapping']).issubset(fields):
            raise ValueError('配置格式不支持')
        mapping = profile['mapping']
    else:
        if Path(args['save_config']).exists():
            raise ValueError('配置文件已存在，请复用或另起名称')
        mapping = {k: args[k] for k in fields if args[k]}
    cli = [args['node'] or 'node', str(ROOT/'reconcile.cjs'), '--left', args['left'], '--right', args['right'], '--python', sys.executable]
    for key, value in mapping.items():
        if key == 'single_key_confirmed':
            if value is True: cli.append('--single-key-confirmed')
        elif isinstance(value, str):
            cli += ['--'+key.replace('_','-'), value]
        else:
            raise ValueError('配置字段类型无效')
    def execute(extra):
        result = subprocess.run(cli + extra, capture_output=True, text=True, timeout=60)
        if result.returncode:
            raise ValueError(result.stderr.strip())
        return json.loads(result.stdout)
    inspection = execute(['--inspect'])
    headers = {k: inspection[k]['headers'] for k in ['left', 'right']}
    if args['config'] and headers != profile['headers']:
        raise ValueError('表头与保存配置不同，请先确认字段变化并建立新配置')
    if not args['config']:
        profile = {'version': 1, 'mapping': mapping, 'headers': headers}
    out = Path(args['out'])
    if out.exists():
        raise ValueError('结果目录已存在，请使用新目录')
    # Stage outputs; invalid previous reviews never leave a partial success directory.
    with tempfile.TemporaryDirectory(dir=out.parent) as staging:
        stage = Path(staging)/'result'
        summary = execute(['--out', str(stage)])
        changes = build_workbook(read_json(stage/'核对记录.json'), summary, profile, args['previous'], stage/'复核工作簿.xlsx')
        manifest = read_json(stage/'运行清单.json')
        manifest['workflowVersion'] = VERSION
        manifest['profileSHA256'] = digest(profile)
        manifest['reviewChanges'] = changes
        if args['previous']:
            manifest['previousReviewSHA256'] = hashlib.sha256(Path(args['previous']).read_bytes()).hexdigest()
        manifest['outputs'].append({'file': '复核工作簿.xlsx', 'sha256': hashlib.sha256((stage/'复核工作簿.xlsx').read_bytes()).hexdigest()})
        (stage/'运行清单.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
        if args['save_config']:
            with Path(args['save_config']).open('x', encoding='utf-8') as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
        stage.rename(out)
    print(json.dumps({'workbook': str(out/'复核工作簿.xlsx'), 'changes': changes}, ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print('处理失败：'+str(error), file=sys.stderr)
        sys.exit(1)
