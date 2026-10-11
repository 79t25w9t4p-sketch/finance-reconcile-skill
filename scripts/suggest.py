#!/usr/bin/env python3
"""Propose field mappings from headers only; never guess between candidates."""
import argparse
import json
from pathlib import Path
import subprocess
import sys
from reconcile import fingerprint, load_table

ALIASES = {
    'reference': {'编号', '单号', '单据号', '单据编号', '业务单号', '订单号', '订单编号', '送货单号', '送货单编号', '供方送货单号', '供应商送货单号', '发票号', '发票号码', '凭证号', '凭证编号', '流水号', '交易流水号', '交易编号', 'reference', 'reference_id', 'invoice_id', 'order_id'},
    'amount': {'金额', '余额', '含税金额', '含税总额', '未税金额', '不含税金额', '交易金额', '收款金额', '付款金额', '回款金额', '应收金额', '应付金额', '未收金额', '未付金额', '借方金额', '贷方金额', 'amount', 'balance', 'paid_amount'},
    'party_code': {'客户编码', '客户编号', '供应商编码', '供应商编号', '往来单位编码', '往来编码', 'customer_id', 'vendor_id'},
    'party_name': {'客户', '客户名称', '供应商', '供应商名称', '往来单位', '往来单位名称', '对方户名', 'customer', 'customer_name', 'vendor', 'vendor_name'},
    'company': {'公司', '公司名称', '核算主体', '本方公司', 'company', 'company_name'},
}


def propose(left, right):
    def candidates(headers):
        return {kind: [name for name in headers if name.strip().lower() in aliases]
                for kind, aliases in ALIASES.items()}
    sides = {'left': candidates(left), 'right': candidates(right)}
    issues = []
    mapping = {}
    for side, fields in sides.items():
        for kind in ['reference', 'amount']:
            count = len(fields[kind])
            if count != 1:
                issues.append({'side': side, 'field': kind, 'candidates': fields[kind],
                               'reason': '多个候选，请确认用途' if count else '未识别到常见列名，请指定'})
        if len(fields['amount']) == 1:
            mapping[side + '_amount'] = fields['amount'][0]
    contexts = []
    # Never pair a customer code with a customer name, or ignore one-sided scopes.
    for kind in ['company']:
        counts = [len(sides[s][kind]) for s in ['left', 'right']]
        if counts == [1, 1]:
            contexts.append(kind)
        elif any(counts):
            issues.append({'field': kind, 'reason': '公司范围字段不对应，需要确认'})
    party_present = any(sides[s][k] for s in sides for k in ['party_code', 'party_name'])
    if all(len(sides[s]['party_code']) == 1 for s in sides):
        contexts.append('party_code')
    elif all(len(sides[s]['party_name']) == 1 for s in sides):
        contexts.append('party_name')
    elif party_present:
        issues.append({'field': 'party', 'reason': '客户或供应商字段不对应，需要确认编码/名称口径'})
    for side, fields in sides.items():
        if len(fields['reference']) == 1:
            mapping[side + '_keys'] = json.dumps([fields[k][0] for k in contexts] + fields['reference'], ensure_ascii=False)
    return {
        'status': 'needs_field_choice' if issues else 'ready_for_confirmation',
        'mapping': mapping,
        'candidates': sides,
        'questions': issues,
        'confirm': ['两表期间、币种、金额含义和明细粒度一致', '建议列确实对应同一业务；无客户/公司字段时，确认编号在本次范围内唯一'],
        'note': '只依据表头提出建议，未读取明细到模型、未执行核对或保存配置。',
    }


def inspect_input(file, sheet=None, header_row=None):
    """Find an unambiguous XLSX layout before the strict table reader runs."""
    path=Path(file)
    before=fingerprint(path,10*1024*1024)
    layout=None
    if path.suffix.lower()=='.xlsx' and not (sheet and header_row):
        options={'sheet':sheet,'headerRow':header_row,
                 'reference':sorted(ALIASES['reference']),'amount':sorted(ALIASES['amount'])}
        result=subprocess.run([sys.executable,'-X','utf8',str(Path(__file__).with_name('xlsx_read.py')),
                               '--layout',str(path),json.dumps(options,ensure_ascii=False)],
                              capture_output=True,text=True,encoding='utf-8',timeout=30)
        if result.returncode: raise ValueError(result.stderr.strip())
        layout=json.loads(result.stdout)
        if len(layout['candidates'])==1:
            selected=layout['candidates'][0]
            sheet,header_row=selected['sheet'],str(selected['headerRow'])
        else:
            if fingerprint(path,10*1024*1024)!=before: raise ValueError('检查过程中输入文件发生变化')
            return {'file':path.name,**layout},None
    table,metadata=load_table(path,sheet,header_row or '1')
    if metadata['sha256']!=before: raise ValueError('检查过程中输入文件发生变化')
    location={}
    if metadata['sheet'] is not None:
        location={'sheet':metadata['sheet'],'header_row':str(metadata['headerRow'])}
    return {'file':path.name,'sheet':metadata['sheet'],'headerRow':metadata['headerRow'],
            'headers':table['headers'],'rows':len(table['rows'])},location


def main():
    parser = argparse.ArgumentParser(description='识别常见编号、金额、客户和公司列，提出核对建议')
    for option in ['left', 'right']:
        parser.add_argument('--'+option, required=True)
    for option in ['left-sheet', 'right-sheet', 'left-header-row', 'right-header-row']:
        parser.add_argument('--'+option)
    args = vars(parser.parse_args())
    inspection = {}
    locations={};choices=[]
    for side in ['left', 'right']:
        inspection[side],locations[side]=inspect_input(args[side],args[side+'_sheet'],args[side+'_header_row'])
        if locations[side] is None:
            choices.append({'side':side,'field':'table','candidates':inspection[side]['candidates'],
                            'reason':'多个明细页/表头候选，请确认位置' if inspection[side]['candidates'] else
                                     '未识别到包含常见编号和金额列的表头，请指定明细页及表头行；自定义列名可手动指定'})
    if choices:
        print(json.dumps({'status':'needs_table_choice','questions':choices,'inputOverview':inspection,
                          'note':'仅检查工作表名称和表头候选，未执行核对。'},ensure_ascii=False,indent=2))
        return
    proposal = propose(inspection['left']['headers'], inspection['right']['headers'])
    for side,location in locations.items():
        proposal['mapping'].update({side+'_'+key:value for key,value in location.items()})
    proposal['inputOverview'] = inspection
    print(json.dumps(proposal, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print('字段识别失败：'+str(error), file=sys.stderr)
        sys.exit(1)
