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
        self.assertGreaterEqual(rate, 0.9)


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
