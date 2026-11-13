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

    def test_first_device_wins(self):
        self.assertEqual(T.topics_of('iPhone 與 Mac 本週零搭售增加'), ['zero:iPhone'])

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


class WeekTest(unittest.TestCase):
    """同一家店跨週（2026-11-13 加）：使用者反映店的這週與上週評語過於類似。"""
    LAST = 'iPhone 本週銷售由 48 隻降至 24 隻，AC+ 搭售率由 20.8% 提升至 41.7%，保護殼也賣出 50 個，請店長追蹤預約顧客，增加結單數量'

    def test_same_store_same_skeleton(self):
        now = 'iPhone 本週銷售由 24 隻增加至 36 隻，AC+ 搭售率由 41.7% 提升至 44.4%，保護殼也賣出 31 個，請店長追蹤報價顧客，穩定結單數量'
        res = T.week_checks({'微風': [now, '']}, {'微風': [self.LAST, '']})
        self.assertIn('❌', statuses(res, '同店跟上週太像'))

    def test_third_week_same_topic(self):
        now = '結帳前說明保固後，iPhone AC+ 搭售率回到 44%，是三週來最高，請店長把做法寫成交機檢查表'
        res = T.week_checks({'微風': [now, '']}, {'微風': [self.LAST, '']}, {'微風': ['iPhone AC+ 搭售率僅 20.8%，請店長逐筆了解', '']})
        self.assertIn('⚠️', statuses(res, '同店連三週同主題'))
        res = T.week_checks({'微風': [now, '']}, {'微風': [self.LAST, '']}, {})
        self.assertEqual(statuses(res, '同店連三週同主題'), [])
        self.assertIn('⚠️', statuses(res, '同店連兩週同主題'))

    def test_opening_concentrated(self):
        c = {f'店{i}': [f'本週 iPhone 銷售 {i + 20} 隻，保貼增加，請店長檢視第 {i} 項', ''] for i in range(5)}
        c['店5'] = ['抵用券兌換率偏低，請店長抽考', '']
        self.assertIn('⚠️', statuses(T.week_checks(c), '開頭寫法太集中'))
        self.assertEqual(T.opening_type('上週提醒 AirPods 零搭售，本週仍有 18 組'), '接續上週')
        self.assertEqual(T.opening_type('本週 3PP 完成 35.1 萬'), '金額')

    def test_different_store_not_flagged(self):
        res = T.week_checks({'士林': [self.LAST, '']}, {'微風': [self.LAST, '']})
        self.assertEqual(res, [])


class SpeechTest(unittest.TestCase):
    LAST = '# 北一區 27Q1W06 講稿（11/1～11/7）\n\n## 完整版\n\n本週北一區總營業額完成 2,442 萬，較前週增加約 172 萬，成長 7.6%。門市部分…'

    def test_same_structure_as_last_week(self):
        now = '# 北一區 27Q1W07 講稿\n\n## 完整版\n\n本週北一區總營業額完成 2,310 萬，較前週減少約 132 萬，下滑 5.4%。'
        res = T.speech_checks(now, self.LAST)
        self.assertIn('❌', statuses(res, '講稿結構跟上週一樣'))      # 兩週都沒標＝都是數字總覽
        self.assertIn('⚠️', statuses(res, '講稿沒標結構'))
        self.assertIn('⚠️', statuses(res, '講稿開場句跟上週同骨架'))

    def test_new_structure_passes(self):
        now = '# 北一區 27Q1W07 講稿（結構：門市故事）\n\n## 完整版\n\n這週想先講阿波羅，CPU 16 台裡有 13 台沒有搭配件。'
        self.assertEqual(T.speech_checks(now, self.LAST), [])


if __name__ == '__main__':
    unittest.main()
