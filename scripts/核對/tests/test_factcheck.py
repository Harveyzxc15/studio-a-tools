# -*- coding: utf-8 -*-
"""核對工具回歸測試：python3 -m unittest discover -s tests（在 scripts/核對 底下跑）

黃金樣本＝Bob 實際寄出的定稿＋當期事實檔（tests/golden/）：
  - 定稿要能全部對回（核對工具不能冤枉對的數字）
  - 定稿數字故意改錯，要被抓出來（核對工具不能放過錯的數字）
改了 factcheck.py／*_facts.py 一定要跑，退出碼 0 才算沒改壞。
"""
import json
import os
import re
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import factcheck as fc

G = os.path.join(HERE, 'golden')


def load(*names):
    facts, checks = [], []
    for n in names:
        with open(os.path.join(G, n), encoding='utf-8') as f:
            d = json.load(f)
        facts += d['facts']
        checks += d.get('meta', {}).get('checks', [])
    return facts, checks


def text_of(name):
    with open(os.path.join(G, name), encoding='utf-8') as f:
        raw = f.read()
    if name.endswith('.json'):
        return [f'{s}：{p}' for s, ps in json.loads(raw).items() for p in ps]
    return [raw.strip()]


CASES = [  # (事實檔, 定稿, profile, 允許 ⚠️ 數)
    (('daily_20261031_facts.json', 'daily_20261031_week.json'), 'daily_20261031_final.txt', fc.N1_STORES, 0),
    (('daily_20261021_facts.json',), 'daily_20261021_final.txt', fc.N1_STORES, 0),
    (('weekly_27Q1W02_facts.json',), 'weekly_27Q1W02_final.json', fc.N1_STORES, 0),
    (('weekly_27Q1W03_facts.json',), 'weekly_27Q1W03_final.json', fc.N1_STORES, 0),
    # 北二 Yuni：「保貼、保護殼佔比 69%、59%」並列寫法，69% 會是 ⚠️（真的要人看），允許 1 個
    (('weekly_n2_27Q1W03_facts.json',), 'weekly_n2_27Q1W03_final.json', fc.N2_STORES, 1),
    # 北一 W04：微風「兌換筆數較前週增加 1 筆」的上週值只在上週評語文字裡、算不出差額 → 允許 1
    (('weekly_27Q1W04_facts.json',), 'weekly_27Q1W04_final.json', fc.N1_STORES, 1),
    # 北二 W04：西門 Duo 段用新品開賣追蹤事實檔；「預約到貨 16 台」是 Yuni 的現場資訊（本來就查不到）、
    # 永和「分別提升至 90%、100%」並列寫法 → 各允許 1
    (('weekly_n2_27Q1W04_facts.json', 'weekly_n2_27Q1W04_duo.json'), 'weekly_n2_27Q1W04_final.json', fc.N2_STORES, 1),
]


def mutate(tok):
    """把一個數字改成「看起來合理但錯」的值，保留原本的格式與單位。"""
    x, d, u = tok['x'], tok['dec'], tok['unit'] or ''
    if u in ('%', '％'):
        y = x + (1.3 if d else 3)
    elif u in ('萬', '億'):
        y = x * 1.08 + (10 ** -d if d else 1)
    elif tok['comma'] or u == '元':
        y = round(x * 1.07)
    else:
        y = x + (3 if x >= 10 else 2)
    if tok['comma']:
        s = f'{y:,.{d}f}'
    else:
        s = f'{y:.{d}f}'
    return s + (' ' if ' ' in tok['text'] else '') + u


class GoldenPass(unittest.TestCase):
    def test_finals_all_traceable(self):
        for files, final, stores, max_warn in CASES:
            facts, _ = load(*files)
            for text in text_of(final):
                res = fc.check_numbers(text, facts, stores)
                bad = [(r['tok']['text'], r['status'], r['why']) for r in res if r['status'] not in ('✅', '🎯')]
                with self.subTest(final=final, text=text[:20]):
                    self.assertLessEqual(len(bad), max_warn, bad)


class Mutation(unittest.TestCase):
    def test_wrong_numbers_are_caught(self):
        total = caught = 0
        missed = []
        for files, final, stores, _ in CASES:
            facts, _ = load(*files)
            idx_facts = facts
            for text in text_of(final):
                for tok in fc.extract_numbers(text):
                    wrong = mutate(tok)
                    mutated = text[:tok['pos']] + wrong + text[tok['end']:]
                    res = fc.check_numbers(mutated, idx_facts, stores)
                    r = next((r for r in res if r['tok']['pos'] == tok['pos']), None)
                    total += 1
                    if r is None or r['status'] != '✅':
                        caught += 1
                    else:
                        missed.append(f"{final} {tok['text']}→{wrong}（{r['why']}）")
        rate = caught / total
        print(f'\n改錯數字抓出率 {caught}/{total} = {rate:.1%}')
        for m in missed:
            print('  漏抓：', m)
        # 底線＝目前實測（2026-11-03 加入 W04 後：原 5 樣本 90.7%、W04 85.3%、合計 88.7%）。
        # 漏抓大多是改錯後剛好等於同句另一個真實的小數字（例「由 3 台增加至 5 台」的 3 改成 5），
        # 只能靠期間語意判斷，試過「由／至」規則沒有改善。改了規則不可低於這條線。
        self.assertGreaterEqual(rate, 0.885)


class Untraceable(unittest.TestCase):
    def test_numbers_outside_facts_fail(self):
        # 10/26 定稿的 Watch 配件件數、金額不在日報 Excel → 一定要 ❌（要另外查、用 check.py add 登記來源）
        facts, _ = load('daily_20261026_facts.json')
        res = fc.check_numbers(text_of('daily_20261026_final.txt')[0], facts, fc.N1_STORES)
        bad = {r['tok']['text'] for r in res if r['status'] == '❌'}
        self.assertIn('571 件', bad)
        self.assertIn('43.6 萬', bad)

    def test_empty_attachment_flagged(self):
        # 10/22 士林附件只有公式沒值：一定要擋下來，不能當成「六店都在」
        _, checks = load('daily_20261022_facts.json')
        self.assertTrue(any(c[0] == '❌' and '士林' in c[2] for c in checks), checks)


class YoYDiff(unittest.TestCase):
    def test_month_vs_last_year_diff(self):
        # 11/17 北二自動盲測 B：本月累計 vs 去年同期（週報產生器引擎）的差額被判「找不到」而整份剔除
        F = []
        for s, a, b in (('', 948, 1477), ('板橋遠百', 231, 352), ('西門', 246, 347)):
            g = f'{s}|iPhone台數'
            F += [fc.fact(a, 'cnt', s, f'{s or "北二區"} iPhone台數 本月累計', g, '週報產生器引擎', '本月累計'),
                  fc.fact(b, 'cnt', s, f'{s or "北二區"} iPhone台數 去年同期', g, '週報產生器引擎', '去年同期')]
        text = ('北二區 iPhone 銷售 948 支，較去年同期 1,477 支減少 529 支，'
                '門市以板橋遠百減少 121 支、西門減少 101 支差異最大')
        res = fc.check_numbers(text, F, fc.N2_STORES)
        self.assertEqual({r['tok']['text']: r['status'] for r in res if r['status'] != '✅'}, {})
        got = {r['tok']['text']: r['why'] for r in res}
        self.assertIn('本月累計 vs', got['529 支'])
        # 差額一定要寫在「減少／增加」後面才算；同一店的差額不能拿去對別店
        res = fc.check_numbers('西門 iPhone 銷售 121 支', F, fc.N2_STORES)
        self.assertEqual(res[0]['status'], '❌')


class ListedDiff(unittest.TestCase):
    def test_diff_word_carries_over_comma_list(self):
        # 11/17 北二試跑 A：「分別成長 5.6%(約 77 萬)、12.4%(約 199.5 萬)」第二個差額在「、」後面，看不到「成長」而被判找不到
        F = []
        for s, a, b in (('裕隆城', 14_500_000, 13_730_000), ('板橋遠百', 18_105_251, 16_109_962)):
            F += [fc.fact(a, 'amt', s, f'{s} 總營業額 月累', f'{s}|總營業額', '門市日報', '月累'),
                  fc.fact(b, 'amt', s, f'{s} 總營業額 上月同期', f'{s}|總營業額', '門市日報', '上月同期')]
        text = '門市以裕隆城、板橋遠百進度較佳，兩店也較上月同期分別成長約 77 萬、約 199.5 萬'
        res = fc.check_numbers(text, F, fc.N2_STORES)
        got = {r['tok']['text']: r['status'] for r in res}
        self.assertEqual(got['199.5 萬'], '✅')
        self.assertNotEqual(got['77 萬'], '❌')          # 「兩店分別」對哪店本來就只給 ⚠️，不在這次範圍
        # 句子裡沒有增減字，「、」列舉的數字還是不能拿差額來對
        res = fc.check_numbers('板橋遠百營業額 1,810 萬、199.5 萬', F, fc.N2_STORES)
        self.assertEqual({r['tok']['text']: r['status'] for r in res}['199.5 萬'], '❌')


class BestMatch(unittest.TestCase):
    def test_picks_closest_source(self):
        # 11/1 草稿：同一個數值對得到好幾筆時，核對單要顯示真正的來源
        facts, _ = load('daily_20261101_facts.json')
        res = fc.check_numbers(text_of('daily_20261101_draft.txt')[0], facts, fc.N1_STORES)
        got = {r['tok']['text']: r['why'] for r in res}
        self.assertIn('阿波羅 3PP搭售率 今日', got['10%'])       # 今日 9.985%，不是上月同期 9.94%
        self.assertIn('羅東 3PP搭售率 今日', got['7.4%'])        # 「低於目標」不是目標值
        self.assertIn('北一區 3PP搭售率 今日', got['9.3%'])      # 子句沒講店＝全區，不是前文的高島屋
        self.assertIn('北一區 iPhone保固搭售率 目標', got['20%'])


class Extraction(unittest.TestCase):
    def test_skips_dates_models_codes(self):
        t = '10/21 起 iPhone 18 Pro 開賣，13:59 發現，27Q1W03 第3週，M5 MacBook，10 月整區'
        self.assertEqual([x['text'] for x in fc.extract_numbers(t)], [])

    def test_reads_units(self):
        got = [(x['x'], x['unit']) for x in fc.extract_numbers('完成 2,238 萬，下滑 3.5%，銷售 277 隻')]
        self.assertEqual(got, [(2238, '萬'), (3.5, '%'), (277, '隻')])


class Rules(unittest.TestCase):
    def status(self, text, profile):
        return {(s, name) for s, name, _ in fc.rule_results(text, profile)}

    def test_3pp_care(self):
        self.assertTrue(any(s == '❌' for s, _ in self.status('本週 3PP 與 SA Care 合併 120 萬', 'weekly_n1')))

    def test_n1_monthly_revenue(self):
        self.assertTrue(any('月營業額' in n for _, n in self.status('北一區總營業額達成率 80%，落後', 'daily_n1')))
        self.assertFalse(any('月營業額' in n for _, n in self.status('本週總營業額較前週下滑 3%', 'daily_n1')))

    def test_n2_weekly_length_and_period(self):
        st = self.status('字' * 135, 'weekly_n2')
        self.assertTrue(any('110' in n for _, n in st))
        self.assertTrue(any('。' in n for _, n in st))

    def test_month_end_countdown_warn(self):
        self.assertTrue(any('月底前' in n for _, n in self.status('月底前剩餘 3 天請加強' + '字' * 220, 'daily_n1')))


if __name__ == '__main__':
    unittest.main()
