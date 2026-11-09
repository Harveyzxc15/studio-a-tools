# -*- coding: utf-8 -*-
"""週報評語跨格檢查（topics.py）回歸測試。案例都來自 27Q1W06 實際發生過的重複。"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import topics as T


def statuses(res, name):
    return [st for st, n, _ in res if n == name]


class TopicTest(unittest.TestCase):
    def test_device_carries_over(self):
        t = 'AirPods 本週銷售 34 組，18 組未搭配件，零搭售佔比由 45.8% 增加至 52.9%，請店長逐筆檢視'
        self.assertEqual(T.main_topic(t), 'zero:AirPods')
        self.assertEqual(T.polarity(t, 'zero:AirPods'), '-')

    def test_polarity_directions(self):
        self.assertEqual(T.polarity('iPhone AC+ 本週由 29% 提升至 47.5%，請夥伴維持', 'ac:iPhone'), '+')
        self.assertEqual(T.polarity('AC+ 搭售率由 36.9% 降至 25.2%', 'ac:iPhone'), '=')  # 沒有品類就不算 iPhone
        self.assertEqual(T.polarity('iPhone 零搭售也由 19 隻減少至 12 隻', 'zero:iPhone'), '+')
        self.assertEqual(T.polarity('本週 3PP 完成約 16.6 萬，較前週下滑 32.8%', 'tpp'), '-')

    def test_action_not_topic(self):
        # 要求裡提到的品項不算這格的主題
        self.assertNotIn('case:iPhone', T.topics_of('本週喇叭由 3.1 萬降至 7,900 元，請夥伴交機時推薦保護殼'))


class CrossTest(unittest.TestCase):
    def test_same_problem_three_stores(self):
        c = {'士林': ['AirPods 零搭售本週仍未改善，34 組中有 18 組未搭配件，請店長檢視', '喇叭下滑，請門市邀約試聽'],
             '微風': ['本週 AirPods 銷售 8 組皆未搭配件，請店長確認介紹', 'iPad 銷售 10 台，請夥伴推薦鍵盤'],
             '美麗華': ['AirPods 零搭售佔比由 51.5% 增加至 72.4%，請夥伴帶出掛繩', '原廠筆由 5 支增加至 13 支，請門市整理']}
        self.assertIn('❌', statuses(T.cross_checks(c), '同一個問題超過兩店'))

    def test_same_topic_different_direction_ok(self):
        c = {'士林': ['iPhone AC+ 搭售率由 33.3% 提升至 43.8%，請門市保持', ''],
             '微風': ['iPhone AC+ 由 20% 提升至 44%，請店長追蹤', ''],
             '阿波羅': ['iPhone AC+ 搭售率由 36.9% 降至 25.2%，請店長確認抽考', '']}
        self.assertEqual(statuses(T.cross_checks(c), '同一個問題超過兩店'), [])

    def test_identical_action_two_stores(self):
        a = '請店長 by 人員檢視耳機結單狀況，請夥伴結帳前確認保護殼需求。'
        c = {'永和': ['AirPods 零搭售 10 組，' + a, ''], '板橋遠百': ['AirPods 零搭售 32 組，' + a, '']}
        self.assertIn('❌', statuses(T.cross_checks(c), '要求動作重複'))

    def test_repeat_last_week(self):
        last = {'新莊宏匯': ['', '10月 Watch 零搭售 57 支，請夥伴在手錶結單時主動推薦錶帶與保貼，依顧客穿搭需求提供試戴，降低零搭售比例。']}
        now = {'新莊宏匯': ['CPU 由 8 台增加至 12 台，請夥伴保持',
                         'Watch 零搭售 12 支，請夥伴在手錶結單時主動推薦錶帶與保貼，依顧客穿搭需求提供試戴，降低零搭售比例。']}
        self.assertIn('❌', statuses(T.cross_checks(now, last), '要求動作重複'))

    def test_same_store_same_number(self):
        c = {'花蓮': ['本週 3PP 成長 7.5%，iPad 由 2 台增加至 6 台，請夥伴保持',
                    'ARpedia 開張，iPad 也由 2 台增加至 6 台，請店長安排分享']}
        self.assertIn('❌', statuses(T.cross_checks(c), '同店兩格重提同一數字'))

    def test_source_not_final(self):
        self.assertIn('⚠️', statuses(T.cross_checks({'士林': ['a', 'b']}, {}, '自己上週的簡報'), '上週評語來源'))
        self.assertEqual(statuses(T.cross_checks({'士林': ['a', 'b']}, {}, '定稿信（Bob.Yu@studioa.com.tw）'), '上週評語來源'), [])


if __name__ == '__main__':
    unittest.main()
