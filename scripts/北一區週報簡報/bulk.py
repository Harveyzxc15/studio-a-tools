# -*- coding: utf-8 -*-
"""大單辨識＋夥伴零搭售（寫評語選題用，簡報數字不動）。

  python3 bulk.py 27Q1W06 [--region 北二區] [--source DSS|EPB]

從 EPB 重建本週 3PP 明細（不排除任何客戶），找出「大單」：同一客戶本週同店同品類 ≥ BIG 台。
分兩種：企業戶／通訊業（posvip.occupation_id 88／99；EPB 重建本來就排除，DSS 有時沒排除，
例如 W06 西門 iPhone 21 台、遠百 AirPods 25 台）與其他客戶（例如盤商沒設職業別）。
每店×品類記台數與零搭售台數 → <週次>/bulk.json；tracker.py 用 cut_of() 算「排除大單後」的零搭售率挑問題格。
夥伴零搭售：排除大單後，各店零搭售台數最多的夥伴（EPB 重建，store＋人員彙總對 DSS 驗證 98～100%）。
"""
import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import dss_rebuild_from_epb as m  # noqa: E402

BIG = 5                                   # 同客戶同店同品類本週幾台算大單
CATS = {'CPU': 'CPU', 'iPhone': 'iPhone', 'iPad': 'iPad', 'Watch': 'Watch', 'Airpods': 'AirPods'}


def hosts_of(df):
    """label() 結果 → 主機列：cat（簡報品類名）、zero、q。"""
    h = df[df.type_name.str.endswith('搭售')].copy()
    h['zero'] = h.type_name.str.endswith('零搭售')
    h['cat'] = h.type_name.str.replace('零搭售', '').str.replace('搭售', '').map(CATS)
    h['q'] = pd.to_numeric(h.STK_QTY)
    return h


def analyse(h, excluded, codes):
    """h＝不排除客戶的主機列；excluded＝88／99 客戶 vip_id 集合；codes＝{shop_id: 店名}。
    大單＝同客戶本週同店同品類 ≥ BIG 台。每格記：
      big_n／big_z：不是 88／99 的大單（DSS、EPB 重建都有算進去）
      ex_n／ex_z：88／99 的大單（EPB 重建已排除；DSS 有時有算、有時沒有，tracker 用 epb_n 判斷）
      epb_n：EPB 重建（排除 88／99）的該格主機台數"""
    by_vip = h[h.VIP_ID.notna()].groupby(['SHOP_ID', 'VIP_ID', 'cat']).q.sum()
    big = {k for k, v in by_vip.items() if v >= BIG}
    is_ex = h.VIP_ID.isin(excluded)
    is_big = pd.Series([(r.SHOP_ID, r.VIP_ID, r.cat) in big for r in h.itertuples()], index=h.index, dtype=bool)
    out = {}
    for code, name in codes.items():
        hs = h[h.SHOP_ID == code]
        cats = {}
        for cat, g in hs[is_big[hs.index]].groupby('cat'):
            ex = g[g.VIP_ID.isin(excluded)]
            nx = g[~g.VIP_ID.isin(excluded)]
            vips = [{'vip': vip, 'kind': '企業戶／通訊業' if vip in excluded else '同客戶多台',
                     'n': int(gv.q.sum()), 'z': int(gv[gv.zero].q.sum())} for vip, gv in g.groupby('VIP_ID')]
            cats[cat] = {'big_n': int(nx.q.sum()), 'big_z': int(nx[nx.zero].q.sum()),
                         'ex_n': int(ex.q.sum()), 'ex_z': int(ex[ex.zero].q.sum()),
                         'epb_n': int(hs[(hs.cat == cat) & ~is_ex[hs.index]].q.sum()),
                         'vips': sorted(vips, key=lambda v: -v['n'])}
        # 夥伴零搭售：排除大單與 88／99（這些不是夥伴的問題）
        p = hs[~is_ex[hs.index] & ~is_big[hs.index]]
        people = []
        for emp, g in p.groupby('EMP_ID1'):
            z, n = int(g[g.zero].q.sum()), int(g.q.sum())
            if z > 0:
                by = {c: int(gc[gc.zero].q.sum()) for c, gc in g.groupby('cat') if gc[gc.zero].q.sum() > 0}
                people.append({'emp': emp, 'z': z, 'n': n, 'by': by})
        people.sort(key=lambda x: (-x['z'], x['n']))
        out[name] = {'bulk': cats, 'people': people}
    return out


def cut_of(b, source, cur_n):
    """這格要從簡報數字扣掉的大單 (台數, 零搭售台數)。
    EPB 重建：88／99 本來就排除，只扣其他大單。DSS：88／99 大單只有在 DSS 台數 ≥ EPB 重建＋這批（容差 1）
    時才扣（W06：西門、遠百兩張 88 企業戶 DSS 有算；遠百一批 99 盤商 DSS 已排除）。"""
    n, z = b['big_n'], b['big_z']
    if source == 'DSS' and b['ex_n'] and cur_n is not None and cur_n >= b['epb_n'] + b['ex_n'] - 1:
        n, z = n + b['ex_n'], z + b['ex_z']
    return n, z


def emp_names(ids):
    ids = sorted({i for i in ids if isinstance(i, str) and i})
    if not ids:
        return {}
    df = m.epb(f"SELECT DISTINCT emp_id, name FROM ep_emp WHERE emp_id IN ({','.join(repr(i) for i in ids)})")
    return {r.EMP_ID.strip(): (r.NAME or '').strip() for r in df.itertuples()}


def run(label, region, start, end, source, work):
    import build_deck as B
    B.configure(region)
    lines, occ = m.pull(start.isoformat(), end.isoformat(), m.REGIONS[B.R['dept']])
    excluded = set(occ[occ.OCCUPATION_ID.isin(m.EXCLUDE_OCCUPATION)].VIP_ID)
    h = hosts_of(m.label(lines, occ.iloc[0:0]))        # 不排除任何客戶
    res = analyse(h, excluded, {c: n for c, n, _ in B.STORES})
    names = emp_names(p['emp'] for v in res.values() for p in v['people'])
    for v in res.values():
        for p in v['people']:
            p['name'] = names.get(p['emp'], p['emp'])
    data = {'label': label, 'source': source, 'big': BIG, 'stores': res}
    Path(work, 'bulk.json').write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding='utf-8')
    return data


def load(work):
    p = Path(work) / 'bulk.json'
    return json.loads(p.read_text(encoding='utf-8')) if p.exists() else None


def people_md(data, top=2, min_z=3):
    """各店零搭售最多的夥伴（排除大單），給評語「寫到人」用。"""
    md = ['## 夥伴零搭售（排除大單；EPB 重建）', '',
          f'各店零搭售台數最多的 {top} 位（≥{min_z} 台才列）。評語可寫「請店長針對零搭售較多的夥伴…」，'
          '要不要點名由使用者決定；名字不要寫進評語。', '']
    for s, v in data['stores'].items():
        ps = [p for p in v['people'] if p['z'] >= min_z][:top]
        if ps:
            md.append(f'- {s}：' + '；'.join(
                f'{p["name"]} 零搭售 {p["z"]}/{p["n"]} 台（' + '、'.join(f'{c} {k}' for c, k in p['by'].items()) + '）'
                for p in ps))
    if len(md) == 4:
        md.append('（沒有夥伴零搭售 ≥3 台）')
    return '\n'.join(md) + '\n'


def main():
    import build_deck as B
    ap = argparse.ArgumentParser()
    ap.add_argument('label')
    ap.add_argument('--region', default='北一區')
    ap.add_argument('--source', choices=('DSS', 'EPB'), help='簡報 3PP 用哪個來源（預設看 state.json）')
    ap.add_argument('--end', help='週六日期 YYYY-MM-DD（預設看 state.json）')
    a = ap.parse_args()
    B.configure(a.region)
    work = B.OUT_BASE / a.label
    st = json.loads((work / 'state.json').read_text(encoding='utf-8'))
    end = a.end or st.get('wk_end') or re.search(r'(\d{4}-\d{2}-\d{2})', Path(st['xlsx']).name)[1]
    end = B.date.fromisoformat(end)
    src = a.source or st.get('3pp_source') or 'EPB'
    d = run(a.label, a.region, end - B.timedelta(days=6), end, src, work)
    for s, v in d['stores'].items():
        for c, b in v['bulk'].items():
            print(f'{s} {c}：', {k: b[k] for k in ('big_n', 'big_z', 'ex_n', 'ex_z', 'epb_n')}, [f'{x["kind"]} {x["n"]}' for x in b['vips']])


if __name__ == '__main__':
    main()
