# -*- coding: utf-8 -*-
"""達成率分析事實檔＋雙路對帳。

路 A＝各店門市日報附件（DSS 產，給長官看的那份）；路 B＝EPB 用 日報_推雲端.compute() 獨立重算。
  - 實績：台數、保固套數、成交筆數、舊換新要完全一樣；金額差 ≤0.5% 算對（當天快照時點差），超過就 ❌
  - 目標：當月才拿得到內網目標（內網只留當月）→ 逐店逐項比；過去月份只有日報一個來源，標 ⚠️
  - 已知口徑差異不比：Apple 主機／配件（AirPods 歸類不同）、毛利（日報含 SA Care×0.5）、SA Care 金額（價目表缺 SKU）
「達成率 − 應達成率」（落後／超前幾個百分點）由 factcheck.derive 自動算。

用法：python3 achv_facts.py 20261031 [--region 北二區] [--out 路徑]   （需連 VPN）
輸出：facts.json（含日報全部事實）＋同資料夾 雙路對帳.md
"""
import os
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.dirname(HERE))
import daily_facts as DF
from factcheck import save_facts

# 日報列標籤 → EPB compute 欄位（實績）
ACTUAL = {'零售營業額': ['retail'], '訂金': ['deposit'], '3PP營業額': ['tpp'], '訂單出貨': ['order_ship'],
          'CPU台數': ['cpu'], 'iPad台數': ['ipad'], 'iPhone台數': ['iphone'], 'Watch台數': ['watch'],
          'AirPods台數': ['air'], 'CPU保固套數': ['cpu_w'], 'iPad保固套數': ['ipad_w'],
          'iPhone保固套數': ['iphone_w'], 'Watch保固套數': ['watch_w'], 'AirPods保固套數': ['air_w'],
          '成交筆數': ['txn'], '舊換新台數': ['tradein']}
# 日報列標籤 → 內網目標欄位（相加）
TARGET = {'總營業額': ['col1'], 'Apple主機營業額': ['col4'], 'Apple配件營業額': ['col5'], '3PP營業額': ['col6'],
          '保固營業額': ['col7', 'col23'], '總毛利': ['col8'], 'Apple毛利': ['col9'], '3PP毛利': ['col10'],
          'CPU台數': ['col12', 'col13'], 'CPU保固套數': ['col19', 'col20'], 'iPad台數': ['col14'],
          'iPad保固套數': ['col25', 'col15'], 'iPhone台數': ['col16'], 'iPhone保固套數': ['col26', 'col17'],
          'Watch台數': ['col18'], 'Watch保固套數': ['col27', 'col21'], 'AirPods台數': ['col22'],
          'AirPods保固套數': ['col28']}
AMOUNT_TOL = 0.005


def lookup(facts):
    """{(店, 標籤, 期間): 值}（只取各店的原始事實）。"""
    out = {}
    for f in facts:
        if f.get('p') and f['g'] == f"{f['s']}|{f['l'].split(' ')[1]}" if f['s'] else False:
            out[(f['s'], f['l'].split(' ')[1], f['p'])] = f['v']
    return out


def build(day, region):
    facts, checks = DF.build(day, region)
    val = lookup(facts)
    stores = DF.STORES[region]
    d = date(int(day[:4]), int(day[4:6]), int(day[6:]))
    rows, n_bad = [], 0

    # ---- 實績：日報 vs EPB
    import 日報_推雲端 as D
    epb, _ = D.compute(d)
    for code, name in stores.items():
        if (name, '零售營業額', '月累') not in val:
            continue
        m = epb[code]['mtd']
        for label, keys in ACTUAL.items():
            a, b = val.get((name, label, '月累')), sum(m.get(k, 0) for k in keys)
            if a is None:
                continue
            diff = b - a
            if label.endswith(('台數', '套數', '筆數')):
                ok = '✅' if diff == 0 else '❌'
            else:
                r = abs(diff) / abs(a) if a else (0 if b == 0 else 1)
                ok = '✅' if r == 0 else '⚠️' if r <= AMOUNT_TOL else '❌'
            rows.append((ok, name, label, a, b, diff))
    bad = [r for r in rows if r[0] == '❌']
    warn = [r for r in rows if r[0] == '⚠️']
    if not rows:
        checks.append(('❌', '雙路對帳（實績）', '沒有可對帳的資料（日報附件缺或沒值），不能寫達成率'))
    elif bad:
        checks.append(('❌', '雙路對帳（實績）', f'日報與 EPB 有 {len(bad)} 項對不上：'
                       + '、'.join(f'{r[1]}{r[2]} {r[3]:,.0f} vs {r[4]:,.0f}' for r in bad[:6])))
    else:
        checks.append(('✅', '雙路對帳（實績）', f'{len(rows)} 項：台數／套數／筆數全同，金額差都 ≤{AMOUNT_TOL:.1%}'
                       + (f'（{len(warn)} 項有小差，見雙路對帳.md）' if warn else '')))

    # ---- 目標：日報 vs 內網（內網只留當月）
    trows = []
    if d.strftime('%Y%m') == date.today().strftime('%Y%m'):
        tg = D.intranet_targets(d)
        for code, name in stores.items():
            if code not in tg:
                checks.append(('⚠️', '目標', f'內網沒有 {name} 的目標'))
                continue
            for label, cols in TARGET.items():
                a, b = val.get((name, label, '目標')), sum(tg[code].get(c, 0) for c in cols)
                if a is None:
                    continue
                ok = '✅' if abs(a - b) <= 1 else '❌'
                trows.append((ok, name, label, a, b, b - a))
        tb = [r for r in trows if r[0] == '❌']
        if not trows:
            checks.append(('❌', '雙路對帳（目標）', '沒有可對帳的目標（日報附件缺或內網沒資料）'))
        else:
            checks.append(('❌' if tb else '✅', '雙路對帳（目標）',
                           f'日報目標與內網 {len(trows)} 項' + (f'有 {len(tb)} 項不同：' + '、'.join(
                               f'{r[1]}{r[2]} {r[3]:,.0f} vs {r[4]:,.0f}' for r in tb[:6]) if tb else '全同')))
    else:
        checks.append(('⚠️', '雙路對帳（目標）', f'{d:%Y/%m} 不是當月，內網已沒有該月目標，目標只有日報一個來源'
                       '（要寫目標數字時，另外跟全國目標表／北一區線上月目標表核對）'))

    return facts, checks, rows, trows


def write_md(path, day, region, rows, trows):
    lines = [f'# {region} {day} 雙路對帳', '', '路 A＝門市日報附件；路 B＝EPB 重算（日報_推雲端.compute）', '',
             '## 實績（本月累積）', '', '| | 門市 | 項目 | 日報 | EPB | 差 |', '|---|---|---|---:|---:|---:|']
    lines += [f'| {r[0]} | {r[1]} | {r[2]} | {r[3]:,.0f} | {r[4]:,.0f} | {r[5]:+,.0f} |' for r in rows]
    if trows:
        lines += ['', '## 目標', '', '| | 門市 | 項目 | 日報 | 內網 | 差 |', '|---|---|---|---:|---:|---:|']
        lines += [f'| {r[0]} | {r[1]} | {r[2]} | {r[3]:,.0f} | {r[4]:,.0f} | {r[5]:+,.0f} |' for r in trows]
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    day = args[0]
    region = '北二區' if '--region' in sys.argv and '北二' in sys.argv[sys.argv.index('--region') + 1] else '北一區'
    out = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else \
        os.path.join(DF.OUT_BASE, f'達成率_{region}_{day}', 'facts.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    facts, checks, rows, trows = build(day, region)
    save_facts(out, facts, {'kind': 'achv', 'region': region, 'day': day, 'checks': checks})
    write_md(os.path.join(os.path.dirname(out), '雙路對帳.md'), day, region, rows, trows)
    for c in checks:
        print(*c)
    print(f'事實 {len(facts)} 筆 → {out}')


if __name__ == '__main__':
    main()
