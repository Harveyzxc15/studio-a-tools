# -*- coding: utf-8 -*-
"""大單辨識（bulk.py）＋選題矩陣扣大單（tracker.adjust_bulk）回歸測試"""
import os
import sys
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import bulk as BK  # noqa: E402
import tracker as t  # noqa: E402

t.B.configure('北二區')


def rows(shop, vip, emp, tag, n):
    return [{'SHOP_ID': shop, 'VIP_ID': vip, 'EMP_ID1': emp, 'type_name': tag, 'STK_QTY': '1'} for _ in range(n)]


H = BK.hosts_of(pd.DataFrame(
    rows('050', 'V88', 'E1', 'iPhone零搭售', 21)          # 企業戶大單
    + rows('050', 'V1', 'E2', 'iPhone零搭售', 3)          # 一般客戶 3 台，不算大單
    + rows('050', 'V2', 'E2', 'iPhone搭售', 1)
    + rows('050', 'V3', 'E3', 'Airpods零搭售', 6)         # 沒設職業別的同客戶多台
    + rows('050', None, 'E3', 'Airpods搭售', 2)
    + rows('050', 'Vx', 'E1', 'iPhone配件', 4)))          # 配件列不算主機


class BulkTest(unittest.TestCase):
    def setUp(self):
        self.r = BK.analyse(H, {'V88'}, {'050': '西門'})['西門']

    def test_bulk_kinds(self):
        b = self.r['bulk']
        self.assertEqual((b['iPhone']['ex_n'], b['iPhone']['ex_z'], b['iPhone']['big_n']), (21, 21, 0))
        self.assertEqual(b['iPhone']['epb_n'], 4)                  # EPB 重建（排除 88／99）只剩 4 台
        self.assertEqual((b['AirPods']['big_n'], b['AirPods']['big_z']), (6, 6))

    def test_people_exclude_bulk(self):
        p = {x['emp']: x for x in self.r['people']}
        self.assertNotIn('E1', p)                                  # E1 只有企業戶大單
        self.assertNotIn('E3', p)                                  # E3 的零搭售全是大單
        self.assertEqual((p['E2']['z'], p['E2']['n']), (3, 4))

    def test_cut_dss_only_when_dss_counted_it(self):
        b = self.r['bulk']['iPhone']
        self.assertEqual(BK.cut_of(b, 'DSS', 25), (21, 21))         # DSS 25 台＝4＋21：DSS 有算，要扣
        self.assertEqual(BK.cut_of(b, 'DSS', 4), (0, 0))            # DSS 已排除
        self.assertEqual(BK.cut_of(b, 'EPB', 4), (0, 0))            # EPB 重建本來就排除

    def test_matrix_drops_bulk_flag(self):
        M = {'西門': {'zero:iPhone': {'cur': 24 / 25, 'prev': 0.5, 'cur_n': 25, 'prev_n': 20}}}
        bulk = {'source': 'DSS', 'stores': {'西門': self.r}}
        M2, cut = t.adjust_bulk(M, bulk)
        self.assertAlmostEqual(M2['西門']['zero:iPhone']['cur'], 3 / 4)
        self.assertEqual(cut['西門']['zero:iPhone'], 21)
        self.assertEqual(M['西門']['zero:iPhone']['cur'], 24 / 25)  # 原本的 M 不動

    def test_adjust_skips_when_numbers_dont_fit(self):
        M = {'西門': {'zero:AirPods': {'cur': 0.5, 'prev': 0.5, 'cur_n': 4}}}   # 簡報只有 4 台，扣不了 6 台
        M2, cut = t.adjust_bulk(M, {'source': 'EPB', 'stores': {'西門': self.r}})
        self.assertEqual(M2['西門']['zero:AirPods']['cur'], 0.5)
        self.assertEqual(cut, {})


if __name__ == '__main__':
    unittest.main()
