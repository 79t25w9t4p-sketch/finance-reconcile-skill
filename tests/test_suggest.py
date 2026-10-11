import importlib.util
from pathlib import Path
import json
import os
import subprocess
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('suggest', ROOT/'scripts/suggest.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

class Suggestions(unittest.TestCase):
    def test_delivery_numbers_and_tax_amounts_need_scope_confirmation(self):
        p=module.propose(['送货单号','含税金额'],['供方送货单号','含税金额'])
        self.assertEqual(p['status'],'ready_for_confirmation')
        self.assertEqual(json.loads(p['mapping']['left_keys']),['送货单号'])
        self.assertEqual(json.loads(p['mapping']['right_keys']),['供方送货单号'])
        self.assertIn('两表期间、币种、金额含义和明细粒度一致',p['confirm'])
        ambiguous=module.propose(['送货单号','含税金额','未税金额'],['供方送货单号','含税金额'])
        self.assertEqual(ambiguous['status'],'needs_field_choice')
        self.assertNotIn('left_amount',ambiguous['mapping'])
    def test_different_headers_map_same_roles(self):
        p=module.propose(['客户','订单号','应收金额'],['客户名称','单据编号','回款金额'])
        self.assertEqual(p['status'],'ready_for_confirmation')
        self.assertEqual(json.loads(p['mapping']['left_keys']),['客户','订单号'])
        self.assertEqual(p['mapping']['right_amount'],'回款金额')
    def test_ambiguous_amounts_and_ids_require_choice(self):
        p=module.propose(['订单号','发票号','借方金额','贷方金额'],['单号','金额'])
        self.assertEqual(p['status'],'needs_field_choice')
        self.assertNotIn('left_amount',p['mapping']);self.assertNotIn('left_keys',p['mapping'])
    def test_codes_do_not_match_names(self):
        for left,right in [(['客户编码','编号','金额'],['客户名称','编号','金额']),(['公司','编号','金额'],['编号','金额'])]:
            self.assertEqual(module.propose(left,right)['status'],'needs_field_choice')
    def test_company_and_shared_customer_code_are_kept(self):
        headers=['公司','客户编号','客户名称','单号','金额']
        p=module.propose(headers,headers)
        self.assertEqual(json.loads(p['mapping']['left_keys']),['公司','客户编号','单号'])
    def test_unknown_headers_are_not_fuzzy_guessed(self):
        p=module.propose(['业务备注','未税净额'],['说明','含税总计'])
        self.assertEqual(p['status'],'needs_field_choice')
        self.assertEqual(p['mapping'],{})
    def test_cli_is_read_only_and_does_not_print_detail(self):
        result=subprocess.run([sys.executable,str(ROOT/'scripts/suggest.py'),'--left',str(ROOT/'examples/左表.csv'),'--right',str(ROOT/'examples/右表.csv'),'--node',os.environ.get('NODE','node')],capture_output=True,text=True)
        self.assertEqual(result.returncode,0,result.stderr)
        p=json.loads(result.stdout);self.assertEqual(p['inputOverview']['left']['rows'],7)
        self.assertNotIn('AR001',result.stdout)
        self.assertEqual(json.loads(p['mapping']['left_keys']),['客户','编号'])

if __name__=='__main__': unittest.main()
