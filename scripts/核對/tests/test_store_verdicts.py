"""盲測補充.store_verdicts：逐店判斷門市概況能不能寫。"""
import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import 盲測補充 as BX  # noqa: E402

TEXT = {'004': '今天iPad達成100%的保固搭售率，iPhone 僅達成30%，近期iPhone搭售稍微不理想。會員比數：79%',
        '005': '顧客加購 AppleCare+，夥伴協助配對。會員比數：88%',
        '024': '下午有客人客訴等候太久，店長出面處理',
        '046': ''}
STORES = {'北一區': {'004': '士林', '005': '微風', '024': '美麗華', '046': '阿波羅', '054': '高島屋'}}


def run(facts):
    with mock.patch.object(BX.DF, 'STORES', STORES), \
            mock.patch.object(BX.DF, 'find_files', lambda day, codes: {c: c for c in codes if c in TEXT}), \
            mock.patch.object(BX, 'd43_text', lambda path: TEXT[path]):
        return {v['store']: v for v in BX.store_verdicts('20261111', facts)}


class StoreVerdictTest(unittest.TestCase):
    def test_kinds(self):
        v = run({'北一區': [{'u': 'rate', 's': '士林', 'v': 1.0}, {'u': 'rate', 's': '士林', 'v': 0.304}]})
        self.assertEqual(v['士林']['kind'], '佐證')
        self.assertTrue(v['士林']['reason'].startswith('門市自評：「近期iPhone搭售稍微不理想'))   # 落後的先講
        self.assertNotIn('會員', v['士林']['reason'])
        self.assertNotIn('對不上', v['士林']['reason'])
        self.assertEqual(v['微風']['kind'], '一般')                    # 只有會員比數 % 不算
        self.assertEqual(v['美麗華']['kind'], '事件')
        self.assertEqual(v['阿波羅']['kind'], '沒寫')
        self.assertEqual(v['高島屋']['kind'], '缺')

    def test_mismatch_flagged(self):
        v = run({'北一區': [{'u': 'rate', 's': '士林', 'v': 1.0}]})
        self.assertIn('30% 對不上日報數字', v['士林']['reason'])

    def test_region_without_facts_not_flagged(self):
        v = run({})                                                    # 事實檔有 ❌ 被跳過的區
        self.assertNotIn('對不上', v['士林']['reason'])

    def test_lines(self):
        with mock.patch.object(BX.DF, 'STORES', STORES):
            lines = BX.verdict_lines(list(run({}).values()))
        self.assertTrue(lines[0].startswith('北一 可寫：士林（佐證'))
        self.assertIn('美麗華（事件：提到客訴）', lines[0])
        self.assertEqual(lines[1], '北一 不寫：一般 微風｜沒寫 阿波羅｜缺 高島屋')


if __name__ == '__main__':
    unittest.main()
