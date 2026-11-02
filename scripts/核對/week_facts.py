# -*- coding: utf-8 -*-
"""週對週事實檔（日報週六寫「本週 vs 前週」用）：用週報產生器引擎 multistore_engine.calc_store_metrics
算本週、上週，各店＋區合計。這是 Bob／Yuni 週數字的口徑（10/17、10/31 對過）。

⚠️ 不要用 日報_推雲端.py 的 query_* 相加：日報保固套數含檢測新機 SA Care，
   10/25–31 北一 iPhone 保固會變 45%（產生器 32%），3PP 也不同。

用法：python3 week_facts.py 20261031 [--region 北二區] [--out 路徑]   （日期＝本週週六，需連 VPN）
"""
import importlib.util
import os
import sys
from datetime import date, timedelta
from pathlib import Path

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from factcheck import fact, save_facts
from daily_facts import STORES, OUT_BASE

ENGINE = {'北一區': Path.home() / 'Desktop/north1-weekly-report',
          '北二區': Path.home() / 'Desktop/north2-weekly-report'}
# calc_store_metrics 欄位 → (標籤, 單位)
KEYS = {'total_rev': ('總營業額', 'amt'), 'tpp_excl_sa': ('3PP營業額(未加Care)', 'amt'),
        'sa_rev': ('SA Care營業額', 'amt'), 'acpp_plus': ('AC+營業額', 'amt'), 'coupon_rev': ('禮券營業額', 'amt'),
        'cpu_units': ('CPU台數', 'cnt'), 'ipad_units': ('iPad台數', 'cnt'), 'iphone_units': ('iPhone台數', 'cnt'),
        'watch_units': ('Watch台數', 'cnt'), 'airpods_units': ('AirPods台數', 'cnt'),
        'arpedia_units': ('ARpedia套數', 'cnt'), 'txn_count': ('成交筆數', 'cnt')}
DEV = {'cpu': ('CPU', 'mac'), 'ipad': ('iPad', 'ipad'), 'iphone': ('iPhone', 'iphone'),
       'watch': ('Watch', 'watch'), 'airpods': ('AirPods', 'airpods')}


def load_engine(region):
    spec = importlib.util.spec_from_file_location('multistore_engine', ENGINE[region] / 'multistore_engine.py')
    eng = importlib.util.module_from_spec(spec)
    sys.path.insert(0, str(ENGINE[region]))
    spec.loader.exec_module(eng)
    return eng


def metrics_rows(m):
    """一店一週的事實值：原始欄位＋保固（AC+、保固合計）套數與搭售率、3PP 搭售率。"""
    row = {KEYS[k][0]: (m.get(k) or 0, KEYS[k][1]) for k in KEYS}
    row['保固營業額(AC+SA)'] = ((m['acpp_plus'] or 0) + (m['sa_rev'] or 0), 'amt')
    row['3PP營業額(含SA Care)'] = (m['tpp_rev'] or 0, 'amt')       # 日報週對週寫的 3PP（10/31 Bob 221 萬）
    if m['total_rev']:
        row['3PP搭售率(未加Care)'] = (m['tpp_excl_sa'] / m['total_rev'], 'rate')
        row['3PP搭售率(含SA Care)'] = (m['tpp_rev'] / m['total_rev'], 'rate')
    for k, (name, dev) in DEV.items():
        units = m['cpu_units' if k == 'cpu' else f'{k}_units'] or 0
        acpp, sa = m[f'acpp_{dev}'] or 0, m[f'sa_{k}'] or 0
        row[f'{name}AC+套數'] = (acpp, 'cnt')
        row[f'{name}保固合計套數'] = (acpp + sa, 'cnt')
        if units:
            row[f'{name}AC+搭售率'] = (acpp / units, 'rate')
            row[f'{name}保固合計搭售率'] = ((acpp + sa) / units, 'rate')
    return row


def build(sat, region):
    eng = load_engine(region)
    codes = list(STORES[region])
    this = (sat - timedelta(days=6), sat)
    prev = (this[0] - timedelta(days=7), this[0] - timedelta(days=1))
    df = eng.load_from_epb(prev[0], sat, store_codes=codes)
    try:
        df.attrs['txn'] = eng.load_txn_from_epb(prev[0], sat, codes)
    except Exception:
        pass
    sa_prices = eng.load_sacare_prices(str(ENGINE[region] / 'data' / 'SAcare對應價目表.xlsx'))
    facts = []
    for per, (d1, d2) in (('本週', this), ('上週', prev)):
        for code in codes + [None]:
            who = STORES[region][code] if code else ''
            name = who or region
            for label, (v, u) in metrics_rows(eng.calc_store_metrics(df, d1, d2, code, sa_prices)).items():
                facts.append(fact(v, u, who, f'{name} {label} {per}', g=f'{who}|{label}', src='週報產生器引擎', p=per))
                facts.append(fact(v, u, who, f'{name} {label} {per}', g=f'{label}|{per}', src='週報產生器引擎', p=per))
    return facts, []


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    sat = date(int(args[0][:4]), int(args[0][4:6]), int(args[0][6:]))
    if sat.weekday() != 5:
        sys.exit(f'❌ {sat} 不是週六（週＝週日～週六）')
    region = '北二區' if '--region' in sys.argv and '北二' in sys.argv[sys.argv.index('--region') + 1] else '北一區'
    out = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else \
        os.path.join(OUT_BASE, f'日報_{region}_{args[0]}', 'week_facts.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    facts, checks = build(sat, region)
    save_facts(out, facts, {'kind': 'week', 'region': region, 'sat': args[0], 'checks': checks})
    print(f'事實 {len(facts)} 筆 → {out}')


if __name__ == '__main__':
    main()
