# -*- coding: utf-8 -*-
"""店長週報抽取（manager_reports）回歸測試：python3 -m unittest discover -s tests"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import manager_reports as m  # noqa: E402
import tracker as t  # noqa: E402


class ManagerTest(unittest.TestCase):
    def test_norm_radicals(self):
        self.assertEqual(m.norm('羅東⾨市 ⼈員，⽇期：'), '羅東門市 人員，日期：')

    def test_store_names(self):
        m.B.configure('北二區')
        self.assertEqual(m.store_of('2027Q1W06新西門門市A級.pdf'), '西門')
        self.assertEqual(m.store_of('2027Q1W6宏匯門市B級.pdf'), '新莊宏匯')
        self.assertEqual(m.store_of('2027Q1W6板橋誠品C級.pdf'), '板橋誠品')
        m.B.configure('北一區')
        self.assertEqual(m.store_of('2027Q1W6羅東門市 C級.pdf'), '羅東')
        self.assertEqual(m.store_of('2027Q1W6高島屋門市B級.pdf'), '高島屋')

    def test_flags(self):
        f = m.flags_of('本週有遇到大量採購AirPods5 25台、iPhone 18 Pro 2台\n送禮：CPU＊1、AirPods＊5\n10.MySetup：Mac 82%')
        kinds = [(x['kind'], tuple(x['devices'])) for x in f]
        self.assertIn(('大量採購', ('iPhone', 'AirPods')), kinds)
        self.assertIn(('送禮', ('CPU', 'AirPods')), kinds)
        self.assertEqual(len(f), 2)                       # MySetup 那行不能被標

    def test_prev_label(self):
        self.assertEqual(m.prev_label('27Q1W06'), '27Q1W05')
        self.assertIsNone(m.prev_label('27Q2W01'))

    def test_matrix_note(self):
        mgr = {'板橋遠百': {'flags': [{'kind': '大量採購', 'devices': ['AirPods'], 'text': ''},
                                     {'kind': '外國客', 'devices': [], 'text': ''}]}}
        self.assertEqual(t.mgr_note(mgr, '板橋遠百', 'zero:AirPods'), '（店長：大量採購）')
        self.assertEqual(t.mgr_note(mgr, '板橋遠百', 'zero:CPU'), '')      # 沒寫品類的不亂套
        self.assertEqual(t.mgr_note(mgr, '板橋遠百', 'speaker'), '')


if __name__ == '__main__':
    unittest.main()
