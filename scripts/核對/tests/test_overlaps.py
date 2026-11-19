# -*- coding: utf-8 -*-
"""撞句偵測：兩區草稿跟近期實寄共同的長片語要全部抓出來，改寫時才能一次改乾淨。"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import 盲測補充 as BX


class Overlaps(unittest.TestCase):
    def test_finds_all_phrases_regardless_of_order(self):
        # 11/18 北一 C：第一次只抓到「店長了解 Watch 台數落後原因」，改完才冒出「檢視 Watch 結帳時的保固說明」
        text = '請阿波羅店長了解 Watch 台數落後原因，高島屋這週逐筆檢視 Watch 結帳時的保固說明'
        other = '請西門、板橋誠品店長 By 人員檢視 Watch 結帳時的保固說明，也請永和店長了解 Watch 台數落後原因'
        got = BX.overlaps(text, [other])
        self.assertTrue(any('台數落後原因' in g for g in got), got)
        self.assertTrue(any('結帳時的保固說明' in g for g in got), got)

    def test_short_or_numeric_matches_ignored(self):
        self.assertEqual(BX.overlaps('11 月 AirPods 銷售 545 個', ['11 月 AirPods 銷售 309 個']), [])


if __name__ == '__main__':
    unittest.main()
