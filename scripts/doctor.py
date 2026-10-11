#!/usr/bin/env python3
"""Read-only runtime checks. Never install dependencies or inspect user files."""
import argparse
import json
from pathlib import Path
import sys


def check_environment():
    checks = []
    python_ok = sys.version_info >= (3, 9)
    checks.append({'name': 'Python', 'ok': python_ok,
                   'detail': sys.version.split()[0],
                   'next': '' if python_ok else '需要 Python 3.9 或以上；请让助手检查已有解释器。'})
    try:
        import openpyxl
        version = openpyxl.__version__
        parts = tuple(int(x) for x in version.split('.')[:2])
        excel_ok = (3, 1) <= parts < (4, 0)
        detail = version
    except (ImportError, ValueError, AttributeError):
        excel_ok, detail = False, '当前 Python 缺少可用的 openpyxl（>=3.1,<4）'
    checks.append({'name': 'Excel 读写组件', 'ok': excel_ok, 'detail': detail,
                   'next': '' if excel_ok else '使用运行此检查的同一个 Python，在隔离环境安装 requirements.txt；先说明安装位置。'})
    root = Path(__file__).resolve().parents[1]
    required = ['SKILL.md', 'scripts/workflow.py', 'scripts/reconcile.py',
                'scripts/core.py', 'scripts/xlsx_read.py', 'scripts/suggest.py',
                'requirements.txt', 'examples/month_end_demo.py']
    missing = [name for name in required if not (root/name).is_file()]
    checks.append({'name': '完整技能目录', 'ok': not missing,
                   'detail': '完整' if not missing else '缺少：'+'、'.join(missing),
                   'next': '' if not missing else '重新下载完整目录，不能只复制 SKILL.md。'})
    return {'ready': all(item['ok'] for item in checks),
            'python': sys.executable, 'checks': checks,
            'note': '本次仅检查现有运行环境；未安装软件，未读取业务文件。'}


def main():
    parser = argparse.ArgumentParser(description='检查能否运行对账示例；只检查，不安装')
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args()
    report = check_environment()
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print('可以开始示例核对。' if report['ready'] else '还不能开始；请先处理下面的缺项。')
        for item in report['checks']:
            print(('通过' if item['ok'] else '待处理')+' · '+item['name']+'：'+item['detail'])
            if item['next']:
                print('  下一步：'+item['next'])
        print(report['note'])
    return 0 if report['ready'] else 2


if __name__ == '__main__':
    sys.exit(main())
