# -*- coding: utf-8 -*-
"""追蹤帳本（tracker.update_ledger）回歸測試：python3 -m unittest discover -s tests（在週報工具資料夾跑）"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import tracker as t  # noqa: E402

t.B.configure('北一區')
FIN = {'士林': ['AirPods 零搭售佔比由 40% 增加至 46%，請店長逐筆檢視', 'iPhone AC+ 由 20% 提升至 30%，請夥伴保持']}
M = {'士林': {'zero:AirPods': {'cur': 0.53, 'prev': 0.46}, 'ac:iPhone': {'cur': 0.31, 'prev': 0.30}}}


class LedgerTest(unittest.TestCase):
    def test_verdicts(self):
        led = t.update_ledger({'items': []}, 'W06', 'W05', FIN, M)
        v = {i['key']: (i['history'][-1]['verdict'], i['status']) for i in led['items']}
        self.assertEqual(v['zero:AirPods'], ('變差', 'open'))
        self.assertEqual(v['ac:iPhone'], ('持平', 'closed'))      # 上週是肯定的題目，沒回落就結案

    def test_rerun_same_week_no_duplicates(self):
        led = t.update_ledger({'items': []}, 'W06', 'W05', FIN, M)
        n = len(led['items'])
        led = t.update_ledger(led, 'W06', 'W05', FIN, M)
        self.assertEqual(len(led['items']), n)
        self.assertEqual(len(led['items'][0]['history']), 1)

    def test_stale_final_not_reopened(self):
        led = t.update_ledger({'items': []}, 'W06', 'W05', FIN, M)
        n = len(led['items'])
        led = t.update_ledger(led, 'W07', 'W05', FIN, M)          # W07 沒有新定稿，還是 W05
        self.assertEqual(len(led['items']), n)

    def test_ref_falls_back_to_baseline(self):
        it = {'store': '士林', 'key': 'coupon', 'baseline': 0.5, 'history': [], 'status': 'open', 'actions': [], 'polarity': '-'}
        led = t.update_ledger({'items': [it]}, 'X', 'Y', {}, {'士林': {'coupon': {'cur': 0.3, 'prev': None}}})
        self.assertEqual(led['items'][0]['history'][-1]['ref'], 0.5)

    def test_three_weeks_escalate(self):
        led = {'items': []}
        for w, z in (('W06', 0.55), ('W07', 0.56), ('W08', 0.60)):
            led = t.update_ledger(led, w, 'W05' if w == 'W06' else '', FIN if w == 'W06' else {},
                                  {'士林': {'zero:AirPods': {'cur': z, 'prev': None}}})
        self.assertIn('連續 3 週', t.ledger_md(led, 'W08'))


if __name__ == '__main__':
    unittest.main()
