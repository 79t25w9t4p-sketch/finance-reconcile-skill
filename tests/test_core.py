import sys
from pathlib import Path
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'scripts'))
import core as R


class Core(unittest.TestCase):
    def go(self, left, right):
        return R.reconcile(R.parse(left), R.parse(right), 0, 1, 0, 1)

    def test_csv_quotes_bom_multiline_and_tsv(self):
        self.assertEqual(R.parse('\ufeffid,amount\r\n"a,b","1,200.00"')['rows'], [['a,b', '1,200.00']])
        self.assertEqual(R.parse('id\tamount\nA\t0')['rows'][0][1], '0')
        self.assertEqual(R.parse('id,amount\n"A\nB",1')['rows'][0][0], 'A\nB')
        self.assertEqual(R.parse('id,amount\n"A" ,1')['rows'][0][0], 'A')

    def test_malformed_and_ambiguous_headers(self):
        for text in ['id,id\na,1', 'id,amount\na', 'id,amount\n"a,1',
                     'id,amount\n"a"x,1', 'id,\na,1', 'id,amount\na"b,1']:
            with self.subTest(text=text), self.assertRaises(ValueError):
                R.parse(text)

    def test_exact_decimal_negative_and_invalid_amounts(self):
        for raw, cents in [('1,234.56', 123456), ('(12.30)', -1230), ('0', 0), ('+10.1', 1010)]:
            self.assertEqual(R.money(raw), cents)
        for raw in ['', '1.234', '1,23', '1e3', 'NaN', '￥20', '１２.００']:
            self.assertIsNone(R.money(raw))
        self.assertEqual(R.fmt(-1), '-0.01')

    def test_text_id_order_independent_and_exact_cents(self):
        report = self.go('id,amount\n001,0.30\nB,2', 'code,balance\nB,2.00\n001,0.29')
        rows = {r['id']: r for r in report['result']}
        self.assertEqual(rows['001']['delta'], 1)
        self.assertEqual(rows['B']['status'], '一致')

    def test_duplicates_keep_every_row_without_aggregation(self):
        report = self.go('id,amount\nA,10\nA,20', 'id,amount\nA,30')
        self.assertEqual(report['result'][0]['status'], '重复编号')
        self.assertEqual(len(report['result'][0]['left']), 2)
        self.assertIsNone(report['result'][0]['delta'])
        self.assertEqual(report['leftTotal'], 3000)

    def test_blank_ids_and_invalid_amounts_remain_visible(self):
        report = self.go('id,amount\n,0\nB,', 'id,amount\nC,5')
        self.assertIsNone(report['leftTotal'])
        rows = {r['id']: r for r in report['result']}
        self.assertEqual(rows['B']['status'], '无效金额')
        self.assertEqual(rows['（空编号）']['left'][0]['amount'], 0)

    def test_only_side_records_include_zero(self):
        report = self.go('id,amount\nA,0', 'id,amount\nB,0')
        self.assertEqual([r['status'] for r in report['result']], ['仅左表', '仅右表'])

    def test_independent_column_mapping_and_invalid_mapping(self):
        left, right = R.parse('amount,id\n20,A'), R.parse('code,balance\nA,20')
        self.assertEqual(R.reconcile(left, right, 1, 0, 0, 1)['result'][0]['status'], '一致')
        for keys, amount in [(0, 0), ([1, 1], 0), ([-1], 0), ([True], 0)]:
            with self.assertRaises(ValueError):
                R.reconcile(left, right, keys, amount, 0, 1)

    def test_csv_export_neutralizes_formulas_and_escapes(self):
        report = self.go('id,amount\n=SUM(1),1', 'id,amount\nB,2')
        csv = R.export_csv(report['result'])
        self.assertTrue(csv.startswith('\ufeff'))
        self.assertIn('"\'=SUM(1)"', csv)

    def test_large_amounts_are_exact_without_decimal_context(self):
        self.assertEqual(R.money('9007199254740993.01'), 900719925474099301)
        amount = '9' * 90 + '.01'
        cents = int('9' * 90) * 100 + 1
        self.assertEqual(R.money(amount), cents)
        self.assertEqual(R.fmt(cents, grouping=False), amount)

    def test_composite_keys_distinguish_customers_and_separators(self):
        left = R.parse('客户,编号,金额\n甲,001,100\n乙,001,100\n"a,b",c,1\na,"b,c",2')
        right = R.parse('公司,单号,余额\n甲,001,100\n乙,001,99\n"a,b",c,1\na,"b,c",2')
        report = R.reconcile(left, right, [0, 1], 2, [0, 1], 2)
        self.assertEqual(len(report['result']), 4)
        self.assertEqual(sum(r['status'] == '金额差异' for r in report['result']), 1)
        self.assertFalse(any(r['status'] == '重复编号' for r in report['result']))

    def test_missing_composite_component_and_duplicate_tuples(self):
        left = R.parse('客户,编号,金额\n,001,100\n甲,001,10\n甲,001,20')
        right = R.parse('客户,编号,金额\n甲,001,30')
        report = R.reconcile(left, right, [0, 1], 2, [0, 1], 2)
        self.assertTrue(any(r['status'] == '空编号' for r in report['result']))
        self.assertEqual(len(next(r for r in report['result'] if r['status'] == '重复编号')['left']), 2)

    def test_limits_and_utf8_byte_count(self):
        for text in ['id,amount', 'id,amount\n'+'A,1\n'*5001, ','.join('c'+str(i) for i in range(51))+'\n'+','.join('1' for _ in range(51)), 'id,amount\n'+'中'*700000+',1']:
            with self.assertRaises(ValueError):
                R.parse(text)


if __name__ == '__main__':
    unittest.main()
