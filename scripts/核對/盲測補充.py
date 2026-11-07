# -*- coding: utf-8 -*-
"""日報自動盲測的補充資料與檢查（日報盲測_自動.py 呼叫）：

- store_notes   門市日報附件 D43 門市自己寫的概況 → 事件候選（失竊、客訴、人事、稽核…）
- compare_day   前一天選項 vs 實寄：選題命中紀錄（命中紀錄.json）＋改寫紀錄（語料/改寫紀錄.md）
- hit_summary   近 30 天選題命中率
- overlaps      跟另一區近期實寄／今天選項撞句（共同長片語）
- bulk_orders   單張單同品類 ≥10 台的大單 → 扣除後的搭售率
- month_facts   月初（資料日＝1 日）寫上個月整月概況：上月整月／目標／達成率、禮券、去年同月（產生器口徑）
- other_regions 其他區主管近期日報（語料/其他區日報.txt），當選題範例
"""
import datetime as dt
import difflib
import json
import math
import re
from pathlib import Path

import openpyxl

import daily_facts as DF
import factcheck as fc

HERE = Path(__file__).resolve().parent
CORPUS = HERE / '語料'
OUT_BASE = Path('~/工具中心/輸出/核對').expanduser()
REWRITE = CORPUS / '改寫紀錄.md'
HITS = OUT_BASE / '自動盲測' / '命中紀錄.json'
OTHER = CORPUS / '其他區日報.txt'
WHO = {'北一區': 'Bob', '北二區': 'Yuni'}


def _d(day):
    return dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))


def _tok_fact(tok):
    """抽出的數字 → (值, 單位)，跟 factcheck.candidates 的比法一致。"""
    u, x = tok['unit'], tok['x']
    if u in ('%', '％'):
        return x / 100, 'rate'
    if u in ('萬', '億'):
        return x * (1e4 if u == '萬' else 1e8), 'amt'
    if u == '元' or tok['comma'] or x >= 10000:
        return x, 'amt'
    return x, 'cnt'


# ---------------------------------------------------------------- 1. 門市事件
EVENT_RE = re.compile(r'失竊|遭竊|偷竊|遺失|掉包|報警|警方|警局|備案|客訴|投訴|申訴|糾紛|面試|離職|到職|報到|人力不足|'
                      r'稽核|颱風|停業|停電|漏水|故障|當機|系統異常|受傷|詐騙|偽鈔|消防|地震|缺貨|斷貨|百貨活動|週年慶|快閃')


def d43_text(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = ws['日報'] if '日報' in ws.sheetnames else ws.worksheets[0]
    return '\n'.join(c.value.strip() for r in ws.iter_rows(min_row=40, max_row=60) for c in r
                     if isinstance(c.value, str) and len(c.value.strip()) > 30)


def store_notes(day):
    """{區: (facts, 表格文字)}：各店 D43 原文；有事件關鍵字的標【可能是事件】。原文裡的數字登記成事實。"""
    out = {}
    for region, st in DF.STORES.items():
        files = DF.find_files(day, list(st))
        facts, lines = [], []
        for code, name in st.items():
            if code not in files:
                continue
            try:
                t = re.sub(r'\s+', ' ', d43_text(files[code])).strip()
            except Exception:
                continue
            if not t:
                continue
            hit = sorted(set(EVENT_RE.findall(t)))
            flag = f'【可能是事件：{"、".join(hit)}】' if hit else ''
            lines.append(f'- {name}{flag}：{t[:450]}')
            for tok in fc.extract_numbers(t):
                v, u = _tok_fact(tok)
                facts.append(fc.fact(v, u, name, f'{name} 門市概況原文「{t[max(0, tok["pos"] - 12):tok["end"]]}」',
                                     src=f'{name} 門市日報附件 D43'))
        text = ''
        if lines:
            text = ('# 門市自己寫的今日概況（門市日報 D43 原文；標【可能是事件】的才考慮寫事件題，'
                    '只能整理原文寫到的經過與處理，不能補細節或編原因）\n' + '\n'.join(lines))
        out[region] = (facts, text)
    return out


# ---------------------------------------------------------------- 2./命中率 實寄對比
TOPIC_KW = re.compile(r'iPad mini|AirPods 4|Mac mini|MacBook|iPhone|iPad|CPU|Mac|Watch|AirPods|3PP|喇叭|ARpedia|Duo|'
                      r'週對週|本週|前週|營業額|重點項目|新品|首賣|開賣|MyCoach|稽核|失竊|面試|客訴|月份')


def _kws(t):
    return set(TOPIC_KW.findall(t))


def _clauses(t):
    return [c.strip() for c in re.split(r'[，。；\n]', t) if c.strip()]


def _no_num(t):
    return re.sub(r'[\d.,%\s]', '', t)


def rewrites(opt, sent, limit=14):
    """被採用的選項 → 實寄版：逐子句比對出 改／刪／加（只差數字的不算）。"""
    a, b = _clauses(opt), _clauses(sent)
    sm = difflib.SequenceMatcher(None, [_no_num(x) for x in a], [_no_num(x) for x in b], autojunk=False)
    out = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            continue
        A, B = '，'.join(a[i1:i2]), '，'.join(b[j1:j2])
        if tag == 'replace' and difflib.SequenceMatcher(None, A, B).ratio() >= 0.35:
            out.append(f'- 改：「{A}」→「{B}」')
        else:
            if A:
                out.append(f'- 刪：「{A}」')
            if B:
                out.append(f'- 加：「{B}」')
    return out[:limit]


def compare_day(prev, sent):
    """prev 那天的自動選項 vs 實寄（sent＝{區: 文字}）。記命中紀錄、改寫紀錄；回傳顯示用字串清單。"""
    res = OUT_BASE / '自動盲測' / prev / 'result.json'
    if not res.exists():
        return None
    opts = [o for o in json.loads(res.read_text(encoding='utf-8'))['options'] if o.get('ok')]
    hits = json.loads(HITS.read_text(encoding='utf-8')) if HITS.exists() else {}
    log = REWRITE.read_text(encoding='utf-8') if REWRITE.exists() else ''
    shown = []
    for region in ('北一區', '北二區'):
        if region not in sent:
            continue
        mine = [o for o in opts if o['region'] == region]
        if not mine:
            continue
        s = sent[region]
        scored = []
        for o in mine:
            sim = difflib.SequenceMatcher(None, o['text'], s).ratio()
            same = bool(_kws(o['topic']) & _kws(s[:80])) or sim >= 0.45
            scored.append((same, sim, o))
        scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
        same, sim, best = scored[0]
        tag = '幾乎照用' if sim >= 0.6 else '同題改寫' if same else '換了主題'
        hits.setdefault(prev, {})[region] = {
            'best': best['id'], 'topic': best['topic'], 'sim': round(sim, 3), 'tag': tag, 'hit': tag != '換了主題',
            'a_hit': tag != '換了主題' and best['id'] == 'A', 'n_opts': len(mine), 'sent_head': s[:60]}
        shown.append(f'{region[:2]} 最接近選項 {best["id"]}「{best["topic"]}」相似 {sim:.0%}（{tag}）')
        head = f'### {prev} {region}（{WHO[region]}，選項 {best["id"]}「{best["topic"]}」{tag}）'
        if tag != '換了主題' and head not in log:
            rw = rewrites(best['text'], s)
            if rw:
                log += f'\n{head}\n' + '\n'.join(rw) + '\n'
        elif tag == '換了主題' and head not in log:
            log += f'\n{head}\n- 實寄寫的是：「{s[:120]}…」（我的選項沒有這題）\n'
    HITS.parent.mkdir(parents=True, exist_ok=True)
    HITS.write_text(json.dumps(hits, ensure_ascii=False, indent=1, sort_keys=True), encoding='utf-8')
    if log:
        if not log.startswith('# '):
            log = '# 改寫紀錄（自動：被採用的選項 vs 實寄版；寫稿照這些改寫方向）\n' + log
        REWRITE.write_text(log, encoding='utf-8')
    return shown


def hit_summary(day, n=30):
    if not HITS.exists():
        return ''
    hits = json.loads(HITS.read_text(encoding='utf-8'))
    days = sorted(k for k in hits if k < day)[-n:]
    parts = []
    for region in ('北一區', '北二區'):
        rows = [hits[k][region] for k in days if region in hits[k]]
        if rows:
            a = sum(r['a_hit'] for r in rows)
            h = sum(r['hit'] for r in rows)
            u = sum(r['tag'] == '幾乎照用' for r in rows)
            parts.append(f'{region[:2]} A 命中 {a}/{len(rows)}、任一選項 {h}/{len(rows)}（幾乎照用 {u}）')
    return f'🎯 選題命中率（近 {len(days)} 天有紀錄）：' + '｜'.join(parts) if parts else ''


def rewrite_tail(n=70):
    if not REWRITE.exists():
        return ''
    return '\n'.join(REWRITE.read_text(encoding='utf-8').splitlines()[-n:])


# ---------------------------------------------------------------- 4. 撞句
def overlaps(text, others, n=8):
    """跟 others 共同的長片語（去掉數字標點後 ≥ n 字）。"""
    found = set()
    for o in others:
        sm = difflib.SequenceMatcher(None, text, o, autojunk=False)
        for m in sm.get_matching_blocks():
            seg = text[m.a:m.a + m.size].strip('，、。； ')
            if len(re.sub(r'[\d\s.,%，、。；A-Za-z+]', '', seg)) >= n:
                found.add(seg)
    return sorted(found, key=len, reverse=True)


# ---------------------------------------------------------------- 5. 大單
HOST_CAT = {'4001': 'CPU', '4002': 'CPU', '4004': 'iPhone', '4005': 'iPad', '4006': 'iPad', '4041': 'iPad',
            '4038': 'Watch', '4014': 'AirPods'}
ACPP_KW = {'CPU': ('macbook', 'mac mini', 'imac', 'mac studio'), 'iPhone': ('iphone',), 'iPad': ('ipad',),
           'Watch': ('watch',), 'AirPods': ('airpods',)}


def bulk_orders(day, facts_by_region, epb, base_cond, shop_of, min_qty=10):
    """月累期間單張單同品類主機 ≥ min_qty 台。回傳 {區: (facts, 表格文字)}；用保固圖口徑事實算扣除後搭售率。"""
    d = _d(day)
    m0 = d.replace(day=1)
    cats = ','.join(f"'{c}'" for c in HOST_CAT)
    big = epb(f"SELECT l.shop_id, l.doc_id, MAX(l.doc_date) dd, l.cat4_id, SUM(l.stk_qty) q FROM poslinev_bi l "
              f"WHERE {base_cond(m0, d)} AND l.cat4_id IN ({cats}) AND l.cat3_id IN ('3001','3002') "
              f"GROUP BY l.shop_id, l.doc_id, l.cat4_id HAVING SUM(l.stk_qty)>={min_qty}")
    out = {r: ([], []) for r in DF.STORES}
    for b in big:
        sh = shop_of(b)
        if not sh:
            continue
        region, store = sh
        cat, q = HOST_CAT[b['CAT4_ID'].strip()], int(float(b['Q']))
        lines = epb(f"SELECT l.name, l.stk_qty FROM poslinev_bi l WHERE l.org_id='01' AND l.doc_id='{b['DOC_ID']}' "
                    f"AND l.cat3_id='3032'")
        ac = sum(int(float(x['STK_QTY'])) for x in lines if any(k in x['NAME'].lower() for k in ACPP_KW[cat]))
        facts, txt = out[region]
        src = f'EPB 單號 {b["DOC_ID"]}（同單 cat3 3032 {cat}）'
        facts += [fc.fact(q, 'cnt', store, f'{store} {cat} 大單台數 {b["DD"][5:10]}', src=src),
                  fc.fact(ac, 'cnt', store, f'{store} {cat} 大單同單 AC+ {b["DD"][5:10]}', src=src)]
        line = f'- {store} {b["DD"][5:7]}/{b["DD"][8:10]} 單張單 {cat} {q} 台，同單 AC+ {ac} 套'
        F = {(f['s'], f['l']): f['v'] for f in facts_by_region.get(region, [])}
        for who in (store, ''):
            name = who or region
            n_ = next((v for (s, l), v in F.items() if s.endswith(who) and (who or s == '') and
                       l.endswith(f'{cat} 台數(排認證機) 月累')), None)
            a_ = next((v for (s, l), v in F.items() if s.endswith(who) and (who or s == '') and
                       l.endswith(f'{cat} ACPP+套數 月累')), None)
            if n_ and a_ is not None and n_ - q > 0:
                r = (a_ - ac) / (n_ - q)
                facts.append(fc.fact(r, 'rate', who, f'{name} {cat} ACPP+搭售率 月累(扣除這張大單)',
                                     src=f'({a_:.0f}−{ac})/({n_:.0f}−{q})'))
                line += f'；扣除後{name} {cat} ACPP+ 搭售率 {r:.1%}'
        txt.append(line)
    return {r: (f, ('# 單張單大量訂單（月累期間同品類 ≥10 台；搭售率異常先看這裡，點名門市前用扣除後的數字）\n'
                    + '\n'.join(t)) if t else '') for r, (f, t) in out.items()}


# ---------------------------------------------------------------- 6. 月初整月
def month_facts(day):
    """資料日是 1 日 → 上個月整月。回傳 {區: (facts, 表格文字)}（沒資料的區不在裡面）。"""
    import week_facts as WF
    d = _d(day)
    lm_end = d - dt.timedelta(days=1)
    lm0 = lm_end.replace(day=1)
    M = f'{lm_end.month}月'
    try:
        ly0, ly1 = lm0.replace(year=lm0.year - 1), lm_end.replace(year=lm_end.year - 1)
    except ValueError:
        ly0, ly1 = lm0.replace(year=lm0.year - 1), lm_end.replace(year=lm_end.year - 1, day=28)
    out = {}
    for region in DF.STORES:
        facts, lines = [], [f'# 月初：{M}整月概況（{lm0:%m/%d}～{lm_end:%m/%d}；A 一定寫{M}整月概況）']
        raw, checks = DF.build(lm_end.strftime('%Y%m%d'), region)
        if any(c[0] == '❌' for c in checks):
            continue
        V = {}
        for f in raw:
            p = f.get('p')
            if p in ('月累', '目標', '達成率'):
                newp = {'月累': f'{M}整月', '目標': f'{M}目標', '達成率': f'{M}達成率'}[p]
                f = dict(f, l=f['l'].replace(f' {p}', f' {newp}'), p=newp,
                         g=(f.get('g') or '').replace(p, newp) or None, src=f'{M}底門市日報（{lm_end:%m/%d}）月累')
                facts.append(f)
                V[(f['s'], f['l'].split(' ', 1)[1])] = f['v']
        gen = WF.build_range({M: (lm0, lm_end), f'去年{M}': (ly0, ly1)}, region)
        for f in gen:
            if any(k in f['l'] for k in ('總營業額', '禮券回補', '3PP營業額(含SA Care)', '保固合計搭售率', '台數')):
                facts.append(dict(f, src=f'週報產生器引擎 {f["p"]}'))
        G = {(f['s'], f['l'].split(' ', 1)[1]): f['v'] for f in gen}
        stores = [''] + list(DF.STORES[region].values())
        rows = []
        for s in stores:
            name = s or region
            rev, tgt = V.get((s, f'總營業額 {M}整月')), V.get((s, f'總營業額 {M}目標'))
            rb = G.get((s, f'禮券回補(SA999) {M}'), 0)
            tot, tot_ly = G.get((s, f'總營業額 {M}')), G.get((s, f'總營業額 去年{M}'))
            if rev is None:
                continue
            with_rb = rev + rb
            facts.append(fc.fact(with_rb, 'amt', s, f'{name} 含禮券營業額 {M}', src=f'{M}底日報月累 {rev:,.0f}＋SA999 禮券回補 {rb:,.0f}'))
            cell = f'{name} 營業額 {rev:,.0f}／目標 {tgt or 0:,.0f}／達成 {rev / tgt:.1%}' if tgt else f'{name} 營業額 {rev:,.0f}'
            cell += f'／禮券回補 {rb:,.0f}／含禮券 {with_rb:,.0f}'
            if tgt:
                facts.append(fc.fact(with_rb / tgt, 'rate', s, f'{name} 含禮券達成率 {M}', src=f'{with_rb:,.0f}/{tgt:,.0f}'))
                cell += f'（含禮券達成 {with_rb / tgt:.1%}）'
            if tot and tot_ly:
                facts.append(fc.fact(tot / tot_ly - 1, 'rate', s, f'{name} 總營業額 較去年同月 {M}',
                                     src=f'產生器 {tot:,.0f}/{tot_ly:,.0f}−1'))
                cell += f'／較去年同月 {tot / tot_ly - 1:+.1%}（產生器口徑，含禮券）'
            rows.append(cell)
        lines.append('【營業額】' + '；'.join(rows))
        for k in ('3PP營業額', '3PP搭售率'):
            lines.append(f'【{k}】' + '；'.join(
                f'{s or region} {M} ' + ('/'.join(
                    (f'{V[(s, f"{k} {p}")]:.2%}' if '率' in k else f'{V[(s, f"{k} {p}")]:,.0f}') if (s, f'{k} {p}') in V else '—'
                    for p in (f'{M}整月', f'{M}目標', f'{M}達成率') if not ('率' in k and p == f'{M}達成率')))
                for s in stores))
        for c in ('CPU', 'iPad', 'iPhone', 'Watch', 'AirPods'):
            cells = []
            for s in stores:
                a, t = V.get((s, f'{c}保固搭售率 {M}整月')), V.get((s, f'{c}保固搭售率 {M}目標'))
                ly = G.get((s, f'{c}保固合計搭售率 去年{M}'))
                n, nt = V.get((s, f'{c}台數 {M}整月')), V.get((s, f'{c}台數 {M}目標'))
                cells.append(f'{s or region} 台數 {n or 0:.0f}/{nt or 0:.0f}｜保固 {"—" if a is None else f"{a:.1%}"}'
                             f'/目標 {"—" if t is None else f"{t:.1%}"}/去年 {"—" if ly is None else f"{ly:.1%}"}')
            lines.append(f'【{c}】' + '；'.join(cells))
        lines.append('（營業額、達成率用日報口徑；禮券回補＝SA999 月底回補；去年同月比較用產生器口徑。'
                     '保固「去年」＝產生器保固合計（AC+＋SA Care），去年 Watch／AirPods 多是 SA Care，這兩項不要跟去年比）')
        out[region] = (facts, '\n'.join(lines))
    return out


# ---------------------------------------------------------------- 8. 其他區範例
def other_regions(day, n=12, width=260):
    if not OTHER.exists():
        return ''
    blocks = [b for b in OTHER.read_text(encoding='utf-8').split('## ') if b.strip()]
    rows = []
    for b in blocks:
        head, _, body = b.partition('\n')
        k = head.split()
        if len(k) >= 2 and k[0] < day:
            rows.append((k[0], k[1], re.sub(r'\s+', ' ', body).strip()))
    rows = rows[-n:]
    return '\n'.join(f'- {d[4:6]}/{d[6:]} {r}：{t[:width]}' for d, r, t in rows)
