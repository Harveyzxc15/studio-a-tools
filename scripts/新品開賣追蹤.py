#!/usr/bin/env python3
"""新品開賣追蹤：各店首日／累計台數、金額、AC+ 搭售（講稿、日報的新品主題用）。

    python3 新品開賣追蹤.py --cat6 6347 --acp 'iPhone Duo' --start 2026-10-24
    python3 新品開賣追蹤.py --cat6 6347 --acp 'iPhone Duo' --start 2026-10-24 --end 2026-10-31 --gift-cat5 5611
    python3 新品開賣追蹤.py ... --region 01 --label Duo --facts-out <核對資料夾>/duo_facts.json   # 講稿核對用

口徑（2026-10-26 iPhone Duo 開賣時確立，見記憶 feedback_speech_3pp_care_split）：
  台數／金額  stkmas.cat6_id＝--cat6，TRANS_TYPE A/E/H 淨額（stk_qty 退貨已帶負號），金額＝line_total_net
  AC+        cat6 6525 且品名含 --acp（只算 AC+，不含 SA Care）；搭售率＝AC+ 套數÷台數
  贈品        --gift-cat5：開賣送的配件（Duo 殼＝5611）。同單配件件數會被灌大，所以本表不算零搭售，
              只列贈品件數提醒
輸出：螢幕＋ ~/工具中心/輸出/新品開賣追蹤/<cat6>_<start>_<end>.md
"""
import argparse
import json
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

import dss_rebuild_from_epb as m

STORES = {  # 順序同週報簡報
    '01': [('004', '士林'), ('005', '微風'), ('024', '美麗華'), ('046', '阿波羅'), ('054', '高島屋'), ('057', '羅東')],
    '02': [('009', '永和'), ('025', '板橋誠品'), ('050', '西門'), ('055', '花蓮'), ('063', '板橋遠百'),
           ('064', '新莊宏匯'), ('068', '裕隆城')],
}
REGION_NAME = {'01': '北一區', '02': '北二區'}
OUT = Path.home() / '工具中心' / '輸出' / '新品開賣追蹤'


def pull(a, shops):
    sh = ','.join(f"'{s}'" for s in shops)
    acp = a.acp.replace("'", "''")
    gift = f"or s.cat5_id='{a.gift_cat5}'" if a.gift_cat5 else ''
    df = m.epb(f"""select l.shop_id, to_char(l.doc_date,'YYYY-MM-DD') d,
       case when s.cat6_id='{a.cat6}' then 'unit' when s.cat6_id='6525' then 'acp' else 'gift' end k,
       sum(l.stk_qty) q, sum(l.line_total_net) amt
     from poslinev_bi l join stkmas s on s.stk_id=l.stk_id
     where l.org_id='01' and l.shop_id in ({sh}) and l.trans_type in ('A','E','H')
       and l.doc_date between to_date('{a.start}','YYYY-MM-DD') and to_date('{a.end}','YYYY-MM-DD')
       and (s.cat6_id='{a.cat6}' or (s.cat6_id='6525' and l.name like '%{acp}%') {gift})
     group by l.shop_id, to_char(l.doc_date,'YYYY-MM-DD'),
       case when s.cat6_id='{a.cat6}' then 'unit' when s.cat6_id='6525' then 'acp' else 'gift' end""")
    for c in ('Q', 'AMT'):
        df[c] = pd.to_numeric(df[c])
    return df


def rate(acp, units):
    return f'{acp / units:.1%}' if units else '-'


def region_table(df, reg, a, facts):
    g = df[df.SHOP_ID.isin([c for c, _ in STORES[reg]])]
    q = lambda k, d=None, shop=None: g[(g.K == k) & ((g.D == d) if d else True)
                                       & ((g.SHOP_ID == shop) if shop else True)].Q.sum()
    rows = []
    src = f'新品開賣追蹤.py cat6 {a.cat6} {a.start}~{a.end}'

    def add_facts(who, shop=None):
        u, acp = q('unit', shop=shop), q('acp', shop=shop)
        amt = g[(g.K == 'unit') & ((g.SHOP_ID == shop) if shop else True)].AMT.sum()
        for v, unit, lab in ((q('unit', a.start, shop), 'cnt', '首日台數'), (u, 'cnt', '累計台數'),
                             (amt, 'amt', '累計金額'), (acp, 'cnt', 'AC+ 套數')):
            facts.append(dict(v=float(v), u=unit, s=who, l=f'{a.label} {lab}', g=f'{who}{a.label}{unit}', src=src, p=None))
        if u:
            facts.append(dict(v=acp / u, u='rate', s=who, l=f'{a.label} AC+ 搭售率', g=None, src=src, p=None))

    for code, name in STORES[reg]:
        add_facts(name, code)
        u = q('unit', shop=code)
        rows.append([name, f'{q("unit", a.start, code):.0f}', f'{u:.0f}',
                     f'{g[(g.K == "unit") & (g.SHOP_ID == code)].AMT.sum() / 1e4:,.1f}',
                     f'{q("acp", shop=code):.0f}', rate(q('acp', shop=code), u)]
                    + ([f'{q("gift", shop=code):.0f}'] if a.gift_cat5 else []))
    add_facts('')
    u = q('unit')
    rows.append(['合計', f'{q("unit", a.start):.0f}', f'{u:.0f}', f'{g[g.K == "unit"].AMT.sum() / 1e4:,.1f}',
                 f'{q("acp"):.0f}', rate(q('acp'), u)] + ([f'{q("gift"):.0f}'] if a.gift_cat5 else []))
    hdr = ['門市', f'首日 {a.start[5:]}', f'累計至 {a.end[5:]}', '金額(萬)', 'AC+', 'AC+ 搭售率'] \
        + (['贈品件數'] if a.gift_cat5 else [])
    md = [f'## {REGION_NAME[reg]}', '', '| ' + ' | '.join(hdr) + ' |', '|' + '---|' * len(hdr)]
    md += ['| ' + ' | '.join(r) + ' |' for r in rows]
    # 每日走勢（區合計）
    days = pd.date_range(a.start, a.end).strftime('%Y-%m-%d')
    cum_u = cum_a = 0
    md += ['', '| 日期 | 台數 | 累計台數 | 累計 AC+ 搭售率 |', '|---|---|---|---|']
    for d in days:
        cum_u += q('unit', d)
        cum_a += q('acp', d)
        md.append(f'| {d[5:]} | {q("unit", d):.0f} | {cum_u:.0f} | {rate(cum_a, cum_u)} |')
    return md


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cat6', required=True, help='新品的 cat6 代碼（iPhone Duo＝6347）')
    ap.add_argument('--acp', required=True, help='AC+ 品名關鍵字（cat6 6525 品名含此字才算，例：iPhone Duo）')
    ap.add_argument('--start', required=True, help='開賣日 YYYY-MM-DD')
    ap.add_argument('--end', help='統計到哪天（預設昨天）')
    ap.add_argument('--region', choices=['01', '02', 'all'], default='all', help='01 北一、02 北二')
    ap.add_argument('--gift-cat5', help='開賣贈品配件的 cat5（Duo 殼＝5611），只列件數')
    ap.add_argument('--label', default='新品', help='事實檔裡的名稱（例：Duo）')
    ap.add_argument('--facts-out', help='另存成核對工具的事實檔（check.py run --facts 可多給一份）')
    a = ap.parse_args()
    a.end = a.end or str(date.today() - timedelta(days=1))
    regs = ['01', '02'] if a.region == 'all' else [a.region]
    df = pull(a, [c for r in regs for c, _ in STORES[r]])
    md = [f'# 新品開賣追蹤 cat6 {a.cat6}（AC+ 品名含「{a.acp}」）{a.start} ~ {a.end}', '',
          '口徑：TRANS_TYPE A/E/H 淨額；AC+ 不含 SA Care；金額未稅（line_total_net）。'
          + ('贈品配件會灌大同單配件數，不拿來算零搭售。' if a.gift_cat5 else ''), '']
    facts = []
    for r in regs:
        md += region_table(df, r, a, facts) + ['']
    if a.facts_out:
        Path(a.facts_out).write_text(json.dumps({'meta': {}, 'facts': facts}, ensure_ascii=False, indent=0),
                                     encoding='utf-8')
        print(f'事實檔：{a.facts_out}（{len(facts)} 筆）')
    text = '\n'.join(md)
    OUT.mkdir(parents=True, exist_ok=True)
    f = OUT / f'{a.cat6}_{a.start}_{a.end}.md'
    f.write_text(text, encoding='utf-8')
    print(text)
    print(f'\n已存：{f}')


if __name__ == '__main__':
    main()
