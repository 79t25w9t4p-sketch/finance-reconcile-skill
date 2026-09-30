import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import openpyxl

ROOT = Path(__file__).resolve().parents[1]

class Workflow(unittest.TestCase):
    def test_reuse_review_changes_history_and_schema(self):
        with tempfile.TemporaryDirectory() as folder:
            d=Path(folder);left=d/'left.csv';right=d/'right.csv';config=d/'配置.json'
            left.write_text('编号,金额,说明\nA,10,说明\nB,20,说明\nC,30,说明\nD,40,说明',encoding='utf8')
            right.write_text('编号,金额,说明\nA,9,说明\nB,19,说明\nC,29,说明\nD,39,说明',encoding='utf8')
            def run(out,*extra):
                return subprocess.run([sys.executable,str(ROOT/'scripts/workflow.py'),'--node',os.environ.get('NODE','node'),'--left',str(left),'--right',str(right),'--out',str(d/out),*map(str,extra)],capture_output=True,text=True)
            r=run('first','--save-config',config,'--left-keys','["编号"]','--right-keys','["编号"]','--left-amount','金额','--right-amount','金额');self.assertEqual(r.returncode,0,r.stderr)
            saved=json.loads(config.read_text());self.assertNotIn(str(d),config.read_text())
            previous=d/'first/复核工作簿.xlsx';wb=openpyxl.load_workbook(previous);ws=wb['待处理']
            for row in range(2,6):ws.cell(row,4,'已解决');ws.cell(row,5,'=原始意见')
            # Literal text, not an Excel formula.
            for row in range(2,6):ws.cell(row,5).data_type='s'
            wb.save(previous)
            left.write_text('编号,金额,说明\nB,21,说明\nA,10,说明\nC,30,改变了说明\nE,50,新增',encoding='utf8')
            right.write_text('编号,金额,说明\nE,49,新增\nC,29,说明\nA,9,说明\nB,19,说明',encoding='utf8')
            r=run('second','--config',config,'--previous',previous);self.assertEqual(r.returncode,0,r.stderr)
            wb=openpyxl.load_workbook(d/'second/复核工作簿.xlsx');rows={row[0].value:row for row in list(wb['待处理'])[1:]}
            self.assertEqual(rows['A'][2].value,'未变化');self.assertEqual(rows['A'][3].value,'已解决');self.assertEqual(rows['A'][4].value,'=原始意见');self.assertEqual(rows['A'][4].data_type,'s')
            self.assertEqual(rows['B'][3].value,'需重新复核');self.assertEqual(rows['C'][3].value,'需重新复核');self.assertEqual(rows['E'][2].value,'新增')
            self.assertEqual(wb['历史记录']['A2'].value,'D');self.assertEqual(wb['历史记录']['E2'].value,'=原始意见')
            self.assertTrue(wb['待处理'].auto_filter.ref);self.assertEqual(wb['待处理'].freeze_panes,'F2');self.assertEqual(len(wb['待处理'].data_validations.dataValidation),1)
            right.write_text('编号,余额,说明\nA,9,说明',encoding='utf8')
            r=run('invalid','--config',config);self.assertNotEqual(r.returncode,0);self.assertIn('表头',r.stderr);self.assertFalse((d/'invalid').exists());self.assertEqual(saved,json.loads(config.read_text()))

    def test_invalid_previous_leaves_no_partial_output(self):
        with tempfile.TemporaryDirectory() as folder:
            d=Path(folder);(d/'left.csv').write_text('编号,金额\nA,10');(d/'right.csv').write_text('编号,金额\nA,9');(d/'bad.xlsx').write_text('not a workbook')
            r=subprocess.run([sys.executable,str(ROOT/'scripts/workflow.py'),'--node',os.environ.get('NODE','node'),'--left',str(d/'left.csv'),'--right',str(d/'right.csv'),'--out',str(d/'out'),'--save-config',str(d/'config.json'),'--left-keys','["编号"]','--right-keys','["编号"]','--left-amount','金额','--right-amount','金额','--previous',str(d/'bad.xlsx')],capture_output=True,text=True)
            self.assertNotEqual(r.returncode,0);self.assertFalse((d/'out').exists());self.assertFalse((d/'config.json').exists())

if __name__ == '__main__': unittest.main()
