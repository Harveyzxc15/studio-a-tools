# -*- coding: utf-8 -*-
"""兩區日報「本日營運概況」盲測自動化：每天 12:30（mini launchd）寫昨天的 2～3 個選項，核對 0 ❌ 才推 Discord。

流程：等門市日報附件到齊（信件備份另存的 日報附件/日期/）→ daily_facts 產事實檔
     → 補保固圖口徑（只算 ACPP+）本週／月累／上月同期 → Claude 寫北一（Bob）、北二（Yuni）各 2～3 則
     → check 逐則核對（有 ❌ 讓 Claude 修一次，還有就剔除）→ 推「信件追蹤」頻道。

用法：python3 日報盲測_自動.py [YYYYMMDD，預設昨天] [--dry 不推] [--no-wait 不等晚到的門市信]
      python3 日報盲測_自動.py YYYYMMDD --push-only   # 已產好的那天補推到網站 /daily-blind
輸出：~/工具中心/輸出/核對/自動盲測/<日期>/（prompt、Claude 原始回覆、各則草稿與核對單、result.json）
"""
import datetime as dt
import difflib
import io
import json
import math
import os
import re
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPTS = HERE.parent
ARCHIVE_DAILY = SCRIPTS / '信件追蹤' / 'archive' / '日報附件'
if ARCHIVE_DAILY.exists() and not os.environ.get('DAILY_XLSX_DIR'):
    os.environ['DAILY_XLSX_DIR'] = str(ARCHIVE_DAILY)
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(SCRIPTS))

import openpyxl                      # noqa: E402
import daily_facts as DF             # noqa: E402
import factcheck as fc               # noqa: E402
import 盲測補充 as BX                 # noqa: E402

OUT_BASE = Path(os.environ.get('BLIND_OUT') or '~/工具中心/輸出/核對').expanduser()   # BLIND_OUT：試跑時另存，不蓋正式結果
CORPUS = HERE / '語料'
WEBHOOK = Path('~/.config/studioa/mail_digest_discord_webhook.txt').expanduser()   # 跟每日保固圖同頻道
TOKEN = Path('~/.config/studioa/claude_oauth_token').expanduser()
HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
HUB_TOKEN = Path('~/.config/studioa/epb_push_token').expanduser()
CLAUDE = next((p for p in (Path('~/.local/bin/claude').expanduser(),) if p.exists()), 'claude')
ARCHIVE_PY = '/usr/local/bin/python3' if Path('/usr/local/bin/python3').exists() else sys.executable
DEADLINE = dt.time(13, 20)           # 等晚寄的門市信最晚到幾點
PROFILE = {'北一區': 'daily_n1', '北二區': 'daily_n2'}
WHO = {'北一區': 'Bob（使用者本人）', '北二區': 'Yuni'}
CAT_KEYS = [('mac', 'CPU'), ('iphone', 'iPhone'), ('ipad', 'iPad'), ('watch', 'Watch'), ('airpods', 'AirPods')]
ACPP_TGT = {'CPU': 0.60, 'iPhone': 0.20, 'iPad': 0.45, 'Watch': 0.30, 'AirPods': 0.30}
STORE_ALIAS = {'新店裕隆城': '裕隆城'}


def log(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


# ---------------------------------------------------------------- 資料
def missing_stores(day):
    out = {}
    for region, stores in DF.STORES.items():
        files = DF.find_files(day, list(stores))
        out[region] = [stores[c] for c in stores if c not in files]
    return out


# 門市信箱帳號（寄日報的那個）
STORE_BOX = {'士林': 'shihlin', '微風': 'breeze', '美麗華': 'miramar', '阿波羅': 'apollo', '高島屋': 'dayeh',
             '羅東': 'luodong', '永和': 'yonghe', '板橋誠品': 'banqiaoeslite', '西門': 'ximen', '花蓮': 'hualien',
             '板橋遠百': 'banqiaofe21', '新莊宏匯': 'honhui', '裕隆城': 'yulongcity'}


def mail_status(day, stores):
    """缺附件的店：查信件備份有沒有那店 day 的 Daily report 信 → {店: '沒寄信' | '有寄信但沒附件 HH:MM'}。"""
    msgs = SCRIPTS / '信件追蹤' / 'archive' / 'msgs'
    got = {}
    if msgs.exists():
        cutoff = time.time() - 4 * 86400
        for p in msgs.glob('*.json'):
            try:
                if p.stat().st_mtime < cutoff:
                    continue
                d = json.loads(p.read_text(encoding='utf-8'))
            except Exception:
                continue
            box = d.get('sender', '').split('@')[0]
            if day in d.get('subject', '') and re.search(r'daily ?report', d.get('subject', ''), re.I):
                got.setdefault(box, d.get('received', '')[11:16])
    return {s: (f'有寄信但沒附件 {got[STORE_BOX[s]]}' if STORE_BOX.get(s) in got else '沒寄信') for s in stores}


def check_message(day, miss, first):
    md = f'{day[4:6]}/{day[6:]}'
    lines = [f'📋 **{md} 門市日報檢查**（{dt.datetime.now():%H:%M}）']
    for region, stores in DF.STORES.items():
        lack = miss.get(region, [])
        n = len(stores)
        if not lack:
            lines.append(f'{region[:2]} {n}/{n} ✅')
        else:
            why = mail_status(day, lack)
            lines.append(f'{region[:2]} {n - len(lack)}/{n}：缺 ' + '、'.join(f'{s}（{why[s]}）' for s in lack))
    if any(miss.values()):
        lines.append(f'每 10 分鐘再看一次，最晚 {DEADLINE:%H:%M} 缺的店略過照寫' if first else
                     f'已到 {DEADLINE:%H:%M}，照寫：缺店那區略過缺的店，區合計改用 EPB 口徑（不拿其他店推估）')
    else:
        lines.append('13 店都到齊，開始寫概況')
    return '\n'.join(lines)


def wait_for_files(day, wait, notify=None):
    """等門市日報附件到齊。12:30 第一次檢查一定推一則「門市日報檢查」（到齊／缺哪店、沒寄信還是沒附件）。"""
    first = True
    while True:
        if wait and ARCHIVE_DAILY.exists():
            r = subprocess.run([ARCHIVE_PY, str(SCRIPTS / '信件追蹤' / 'mail_archive.py'), '--imap'],
                               capture_output=True, text=True, timeout=600)
            log('信件備份', (r.stdout.strip().splitlines() or [''])[-1], r.stderr.strip()[-200:])
        miss = missing_stores(day)
        done = not any(miss.values()) or not wait or dt.datetime.now().time() >= DEADLINE
        if notify and (first or (done and any(miss.values()))):
            try:
                notify(check_message(day, miss, first and not done))
            except Exception as e:
                log('門市日報檢查推不出去：', e)
        if done:
            return miss
        log('還沒到：', miss, '→ 10 分鐘後再看')
        first = False
        time.sleep(600)


def acpp_facts(day):
    """保固圖口徑（warranty_report_all.query_period）：只算 ACPP+、排認證機。回傳 {區: (facts, 表格文字)}。"""
    import warranty_report_all as W
    d = dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))
    m0 = d.replace(day=1)
    w0 = d - dt.timedelta(days=(d.weekday() + 1) % 7)          # 週日起
    lm_end = (m0 - dt.timedelta(days=1))
    lm0 = lm_end.replace(day=1)
    lm1 = lm0 + dt.timedelta(days=min(d.day, lm_end.day) - 1)
    periods = {'月累': (m0, d), '上月同期': (lm0, lm1)}
    if w0 != m0:                                               # 跨月那週也要
        periods['本週'] = (w0, d)
    out = {}
    src = 'warranty_report_all.query_period（＝每日保固圖：只算 ACPP+、排認證機）'
    for region, cfg in W.REGIONS.items():
        data = {p: W.query_period(cfg, a, b) for p, (a, b) in periods.items()}
        if sum(v[0] for st in data['月累'].values() for v in st.values()) == 0:
            raise RuntimeError('EPB 主機台數全 0（VPN 可能沒連上）')
        facts, lines = [], []
        rng = {p: f'{a:%m/%d}～{b:%m/%d}' for p, (a, b) in periods.items()}
        lines.append('期間：' + '、'.join(f'{p} {r}' for p, r in rng.items()))
        for key, cat in CAT_KEYS:
            tgt = ACPP_TGT[cat]
            facts.append(fc.fact(tgt, 'rate', '', f'{region} {cat} ACPP+搭售率 目標(保固圖)', src=src + f' 目標 {tgt:.0%}'))
            lines.append(f'【{cat}】目標 {tgt:.0%}｜各店 台數／ACPP+／SACare(檢測新機,不計)／搭售率／缺口')
            for p, per in data.items():
                rows, tot = [], [0, 0, 0]
                for store in cfg['stores'] + [region]:
                    if store == region:
                        n, sc, ap = tot
                    else:
                        n, sc, ap = per.get(store, {}).get(key, (0, 0, 0))
                        tot = [tot[0] + n, tot[1] + sc, tot[2] + ap]
                    s = '' if store == region else STORE_ALIAS.get(store, store)
                    name = s or region
                    gap = max(math.ceil(n * tgt - 1e-9) - ap, 0) if n else 0
                    facts += [fc.fact(n, 'cnt', s, f'{name} {cat} 台數(排認證機) {p}', g=f'ACPP+{cat}台數', src=f'{src} {rng[p]}', p=p),
                              fc.fact(ap, 'cnt', s, f'{name} {cat} ACPP+套數 {p}', g=f'ACPP+{cat}套數', src=f'{src} {rng[p]}', p=p)]
                    if n:
                        facts.append(fc.fact(ap / n, 'rate', s, f'{name} {cat} ACPP+搭售率 {p}', g=f'ACPP+{cat}搭售率',
                                             src=f'{src} {rng[p]}：{ap}/{n}', p=p))
                        if p != '上月同期':
                            facts.append(fc.fact(gap, 'cnt', s, f'{name} {cat} ACPP+缺口 {p}',
                                                 src=f'{src} ⌈{n}×{tgt:.0%}⌉−{ap}'))
                    rate = f'{ap / n:.1%}' if n else '—'
                    rows.append(f'{name} {n}/{ap}/{sc}/{rate}/{"—" if not n else ("達標" if gap == 0 else gap)}')
                lines.append(f'  {p}（{rng[p]}）：' + '；'.join(rows))
        out[region] = (facts, '\n'.join(lines))
    return out


# ---------------------------------------------------------------- 月重點項目／新品首賣
FOCUS_CFG = HERE / '口徑規格' / '月重點項目.json'
ALL_SHOPS = {code: (region, name) for region, st in DF.STORES.items() for code, name in st.items()}
HOST_CAT4 = {'4001': 'mac', '4002': 'mac', '4004': 'iphone', '4005': 'ipad', '4006': 'ipad', '4041': 'ipad',
             '4038': 'watch', '4014': 'airpods'}
ACPP_KW = {'mac': ('macbook', 'mac mini', 'imac', 'mac studio'), 'iphone': ('iphone',), 'ipad': ('ipad',),
           'watch': ('watch',), 'airpods': ('airpods',)}
NEW_DAYS = 7          # 第一次有銷售在近幾天內才算新品
NEW_MIN_UNITS = 2     # 首賣以來 13 店合計至少幾台才算（排除半年沒賣、剛好賣出 1 台的冷門機型）


def model_label(names):
    """EPB 品名 → 口語機型名（iPad mini、MacBook Pro 14…）；認不出就回 None。"""
    n = ' '.join(names).upper()
    for pat, lab in ((r'IPAD MINI', 'iPad mini'), (r'IPAD AIR', 'iPad Air'), (r'IPAD PRO', 'iPad Pro'),
                     (r'\bIPAD\b', 'iPad'), (r'MAC MINI', 'Mac mini'), (r'MAC STUDIO', 'Mac Studio'), (r'IMAC', 'iMac'),
                     (r'\bMBP\b|MACBOOK PRO', 'MacBook Pro'), (r'\bMBA\b|MACBOOK AIR', 'MacBook Air'),
                     (r'IPHONE', 'iPhone'), (r'WATCH', 'Apple Watch'), (r'AIRPODS', 'AirPods')):
        if re.search(pat, n):
            size = re.search(r'(?:MBP|MBA|MACBOOK PRO|MACBOOK AIR|IPAD PRO|IPAD AIR)(?: PXD)? (1[1-6])\b', n)
            chip = re.search(r'/(M\d(?: PRO| MAX)?)/', n)
            return ' '.join(x for x in (lab, size.group(1) if size and 'Mac' in lab else None,
                                        chip.group(1).title().replace('Pro', 'Pro') if chip and 'Mac' in lab else None) if x)
    return None


def _epb_rows(sql):
    import warranty_report_all as W
    return W.epb(sql, max_rows=50000)


def _base_cond(a, b):
    shops = ','.join(f"'{c}'" for c in ALL_SHOPS)
    return (f"l.org_id='01' AND l.shop_id IN ({shops}) AND l.doc_date>=TO_DATE('{a:%Y-%m-%d}','yyyy-mm-dd') "
            f"AND l.doc_date<=TO_DATE('{b:%Y-%m-%d}','yyyy-mm-dd') AND l.trans_type IN ('A','E','H') AND l.cat2_id<>'2029'")


def _shop(r):
    try:
        return ALL_SHOPS.get(f"{int(r['SHOP_ID']):03d}")
    except (KeyError, ValueError):
        return None


def _target_lines(region, name, unit, mtd, today, st_tgt, rg_tgt, src_q, src_t, facts):
    """一個品項（重點項目或新品）各店 月累／今日／目標 → facts＋表格列。"""
    stores = list(DF.STORES[region].values())
    tot_m, tot_d = sum(mtd.get(s, 0) for s in stores), sum(today.get(s, 0) for s in stores)
    tgt = rg_tgt if rg_tgt is not None else (sum(st_tgt.values()) if st_tgt else None)
    rows = []
    for s in stores + [region]:
        m, d = (tot_m, tot_d) if s == region else (mtd.get(s, 0), today.get(s, 0))
        t = tgt if s == region else (st_tgt or {}).get(s)
        k = '' if s == region else s
        facts += [fc.fact(m, 'cnt', k, f'{s} {name} 月累', g=f'{name}台數', src=src_q, p='月累'),
                  fc.fact(d, 'cnt', k, f'{s} {name} 今日', g=f'{name}台數', src=src_q, p='今日')]
        if t is not None:
            facts.append(fc.fact(t, 'cnt', k, f'{s} {name} 目標', g=f'{name}台數', src=src_t, p='目標'))
            if t:
                facts.append(fc.fact(m / t, 'rate', k, f'{s} {name} 達成率', src=f'{m}/{t}'))
                facts.append(fc.fact(max(t - m, 0), 'cnt', k, f'{s} {name} 還差', src=f'{t}−{m}'))
        if s == region or st_tgt or m:
            rows.append(f'{s} {m}/{d}/{"—" if t is None else t}')
    return f'  {name}（{unit}）月累/今日/目標：' + '；'.join(rows)


def focus_facts(day):
    """月重點項目（口徑規格/月重點項目.json）＋近 7 天新品首賣。回傳 {區: (facts, 表格文字)}；沒有就空。"""
    d = dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))
    m0 = d.replace(day=1)
    cfg = json.loads(FOCUS_CFG.read_text(encoding='utf-8')).get(f'{d:%Y-%m}', {}) if FOCUS_CFG.exists() else {}
    out = {r: ([], []) for r in DF.STORES}

    # 1) 月重點項目
    for it in cfg.get('items', []):
        mtd, today = {r: {} for r in DF.STORES}, {r: {} for r in DF.STORES}
        for r in _epb_rows(f"SELECT l.shop_id, l.doc_date, SUM(l.stk_qty) q FROM poslinev_bi l "
                           f"WHERE {_base_cond(m0, d)} AND ({it['where']}) GROUP BY l.shop_id, l.doc_date"):
            sh = _shop(r)
            if not sh:
                continue
            q = int(float(r['Q'] or 0))
            mtd[sh[0]][sh[1]] = mtd[sh[0]].get(sh[1], 0) + q
            if r['DOC_DATE'][:10] == f'{d:%Y-%m-%d}':
                today[sh[0]][sh[1]] = today[sh[0]].get(sh[1], 0) + q
        src_q = f"EPB poslinev_bi {m0:%m/%d}～{d:%m/%d} {it['where']}（A/E/H，排認證機）"
        for region in DF.STORES:
            facts, lines = out[region]
            line = _target_lines(region, it['name'], it.get('unit', '台'), mtd[region], today[region],
                                 (it.get('store_targets') or {}).get(region),
                                 (it.get('region_targets') or {}).get(region),
                                 src_q, f"{d.month}月 KPI 重點項目（月重點項目.json）", facts)
            if not lines:
                lines.append(f'# {d.month} 月重點項目（總部 KPI；月累＝{m0:%m/%d}～{d:%m/%d}）')
            lines.append(line + (f"｜口徑：{it['note']}" if it.get('note') else ''))

    # 2) 新品首賣：近 NEW_DAYS 天才第一次有銷售的主機機型（cat6），往回看 180 天
    look0 = d - dt.timedelta(days=180)
    cat4s = ','.join(f"'{c}'" for c in HOST_CAT4)
    first = {}
    for r in _epb_rows(f"SELECT l.cat6_id, MIN(l.doc_date) f FROM poslinev_bi l WHERE {_base_cond(look0, d)} "
                       f"AND l.cat4_id IN ({cat4s}) AND l.cat3_id IN ('3001','3002') AND l.stk_qty>0 GROUP BY l.cat6_id"):
        f0 = dt.date.fromisoformat(r['F'][:10])
        if f0 > d - dt.timedelta(days=NEW_DAYS):
            first[r['CAT6_ID'].strip()] = f0
    for c6, f0 in first.items():
        rows = _epb_rows(f"SELECT l.shop_id, l.doc_id, l.doc_date, l.cat4_id, l.cat6_id, l.cat3_id, l.name, l.stk_qty "
                         f"FROM poslinev_bi l WHERE {_base_cond(f0, d)} AND l.doc_id IN (SELECT doc_id FROM poslinev_bi "
                         f"WHERE org_id='01' AND cat6_id='{c6}' AND doc_date>=TO_DATE('{f0:%Y-%m-%d}','yyyy-mm-dd') "
                         f"AND doc_date<=TO_DATE('{d:%Y-%m-%d}','yyyy-mm-dd'))")
        cat = next((HOST_CAT4.get(x['CAT4_ID'].strip()) for x in rows if x['CAT6_ID'].strip() == c6), None)
        names = sorted({x['NAME'] for x in rows if x['CAT6_ID'].strip() == c6})
        units, today, acpp = ({r: {} for r in DF.STORES} for _ in range(3))
        per_doc = {}                                  # 單號 → [門市, 新品台數, 同單 AC+, 是否今天]
        for x in rows:
            sh = _shop(x)
            if not sh:
                continue
            q = int(float(x['STK_QTY'] or 0))
            dd = per_doc.setdefault(x['DOC_ID'], [sh, 0, 0, x['DOC_DATE'][:10] == f'{d:%Y-%m-%d}'])
            if x['CAT6_ID'].strip() == c6:
                dd[1] += q
            elif x['CAT3_ID'].strip() == '3032' and cat and any(k in x['NAME'].lower() for k in ACPP_KW[cat]):
                dd[2] += q
        for sh, n_new, n_ac, is_today in per_doc.values():
            units[sh[0]][sh[1]] = units[sh[0]].get(sh[1], 0) + n_new
            if is_today:
                today[sh[0]][sh[1]] = today[sh[0]].get(sh[1], 0) + n_new
            # 同單還買了別台同品類主機時，AC+ 不能全算給新品：最多算到新品台數
            # 銷退單（新品 −1、AC+ −1）照原數扣
            add = min(n_ac, n_new) if n_new > 0 and n_ac > 0 else n_ac
            acpp[sh[0]][sh[1]] = acpp[sh[0]].get(sh[1], 0) + add
        if sum(v for r in units.values() for v in r.values()) < NEW_MIN_UNITS:
            continue
        mt = (cfg.get('model_targets') or {}).get(c6, {})
        label = mt.get('name') or model_label(names) or f'新品 cat6 {c6}'
        src_q = f'EPB poslinev_bi cat6 {c6} {f0:%m/%d}～{d:%m/%d}（A/E/H，排認證機；AC+＝同單 cat3 3032 {cat}）'
        for region in DF.STORES:
            facts, lines = out[region]
            stores = list(DF.STORES[region].values())
            if not sum(units[region].values()):
                continue
            lines.append(f'# 新品首賣：{label}（cat6 {c6}，{f0:%m/%d} 首賣，品名：{"／".join(names[:5])}）')
            lines.append(_target_lines(region, label, '台', units[region], today[region],
                                       (mt.get('store_targets') or {}).get(region), None, src_q,
                                       mt.get('src', 'hub target-setting 分機型目標'), facts)
                         .replace('月累/今日', '首賣以來/今日'))
            ac_rows = []
            for s in stores + [region]:
                n = sum(units[region].values()) if s == region else units[region].get(s, 0)
                a = sum(acpp[region].values()) if s == region else acpp[region].get(s, 0)
                k = '' if s == region else s
                facts.append(fc.fact(a, 'cnt', k, f'{s} {label} AC+套數', g=f'{label}AC+', src=src_q))
                if n:
                    facts.append(fc.fact(a / n, 'rate', k, f'{s} {label} AC+搭售率', src=f'{a}/{n}'))
                    ac_rows.append(f'{s} {n}/{a}/{a / n:.1%}')
            lines.append(f'  {label} 同單 AC+ 台數/套數/搭售率：' + '；'.join(ac_rows))
    return {r: (f, '\n'.join(l)) for r, (f, l) in out.items()}


def daily_table(facts, region):
    """日報口徑（門市日報附件）各店主要數字，給 Claude 看的精簡表。"""
    F = {}
    for f in facts:
        if ' ' in f['l']:
            F.setdefault((f['s'], f['l'].split(' ', 1)[1]), f['v'])
    stores = [''] + list(DF.STORES[region].values())
    pct = lambda x: '—' if x is None else f'{x * 100:.1f}%'
    num = lambda x: '—' if x is None else f'{x:,.0f}'
    L = ['（日報口徑：保固搭售率含 SA Care 檢測新機；月累＝1 日～資料日）']
    for c in ['iPhone', 'iPad', 'CPU', 'Watch', 'AirPods']:
        L.append(f'【{c}】店：台數 今日/月累/目標/上月同期｜保固套數 今日/月累｜搭售率 月累/目標/上月同期')
        for s in stores:
            g = lambda k: F.get((s, k))
            L.append(f"  {s or region}：{num(g(c + '台數 今日'))}/{num(g(c + '台數 月累'))}/{num(g(c + '台數 目標'))}/"
                     f"{num(g(c + '台數 上月同期'))}｜{num(g(c + '保固套數 今日'))}/{num(g(c + '保固套數 月累'))}｜"
                     f"{pct(g(c + '保固搭售率 月累'))}/{pct(g(c + '保固搭售率 目標'))}/{pct(g(c + '保固搭售率 上月同期'))}")
    for k in ['總營業額', '3PP營業額', '保固營業額']:
        L.append(f'【{k}】店：今日/月累/目標/達成率/上月同期')
        for s in stores:
            g = lambda kk: F.get((s, kk))
            L.append(f"  {s or region}：{num(g(k + ' 今日'))}/{num(g(k + ' 月累'))}/{num(g(k + ' 目標'))}/"
                     f"{pct(g(k + ' 達成率'))}/{num(g(k + ' 上月同期'))}")
    L.append('【3PP搭售率】店：月累/目標/上月同期')
    for s in stores:
        g = lambda kk: F.get((s, kk))
        L.append(f"  {s or region}：{pct(g('3PP搭售率 月累'))}/{pct(g('3PP搭售率 目標'))}/{pct(g('3PP搭售率 上月同期'))}")
    return '\n'.join(L)


WEEK_ROWS = [('總營業額(扣禮券回補)', 'amt'), ('禮券回補(SA999)', 'amt'), ('3PP營業額(含SA Care)', 'amt'), ('3PP搭售率(含SA Care,扣回補)', 'rate'),
             ('AC+營業額', 'amt'), ('iPhone台數', 'cnt'), ('iPhone保固合計搭售率', 'rate'), ('iPad台數', 'cnt'),
             ('iPad保固合計搭售率', 'rate'), ('CPU台數', 'cnt'), ('CPU保固合計搭售率', 'rate'), ('Watch台數', 'cnt'),
             ('Watch保固合計搭售率', 'rate'), ('AirPods台數', 'cnt'), ('AirPods保固合計搭售率', 'rate')]


def week_facts(day):
    """週六：週報產生器引擎（week_facts.py）本週 vs 上週。回傳 {區: (facts, 表格文字)}。"""
    import week_facts as WF
    sat = dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))
    out = {}
    for region in DF.STORES:
        facts, _ = WF.build(sat, region)
        V = {(f['s'], f['l'].split(' ', 1)[1]): f['v'] for f in facts}
        stores = [''] + list(DF.STORES[region].values())
        L = [f'（週報產生器口徑＝Bob／Yuni 週數字的口徑；本週 {sat - dt.timedelta(days=6):%m/%d}～{sat:%m/%d}，'
             f'上週 {sat - dt.timedelta(days=13):%m/%d}～{sat - dt.timedelta(days=7):%m/%d}；'
             '保固合計搭售率＝(AC+＋SA Care)÷台數）']
        for label, u in WEEK_ROWS:
            cells = []
            for s in stores:
                a, b = V.get((s, f'{label} 本週')), V.get((s, f'{label} 上週'))
                if a is None and b is None:
                    continue
                fmt = (lambda x: '—' if x is None else f'{x * 100:.1f}%') if u == 'rate' else \
                      (lambda x: '—' if x is None else f'{x:,.0f}')
                cells.append(f'{s or region} {fmt(a)}／{fmt(b)}')
            L.append(f'【{label}】本週／上週：' + '；'.join(cells))
        out[region] = (facts, '\n'.join(L))
    return out


YOY_ROWS = ['3PP營業額(未加Care)', 'CPU台數', 'iPad台數', 'iPhone台數', 'Watch台數', 'AirPods台數']


def yoy_facts(day):
    """本月累計 vs 去年同期（同月 1 日～同一天），週報產生器引擎口徑。回傳 {區: (facts, 表格文字)}。
    Bob 11/10（CPU 台數）、11/11（3PP 金額）實寄用的去年同期就是這組數字（11/11 北一 295.2 萬／305.4 萬逐店對過）。"""
    import week_facts as WF
    d = dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))
    per = {'本月累計': (d.replace(day=1), d), '去年同期': (d.replace(year=d.year - 1, day=1), d.replace(year=d.year - 1))}
    out = {}
    for region in DF.STORES:
        facts = [f for f in WF.build_range(per, region) if any(k in f['l'] for k in YOY_ROWS)]
        V = {(f['s'], f['l'].split(' ', 1)[1]): f['v'] for f in facts}
        L = [f'（週報產生器口徑；本月累計 {per["本月累計"][0]:%m/%d}～{d:%m/%d}，去年同期＝去年同一段日期；'
             '3PP 是未加 SA Care 的金額。只有這裡列的項目可以寫「較去年同期」，保固搭售率不跟去年比；'
             '台數跟日報表、保固圖可能差一兩台，寫去年同期對比時今年和去年都用這裡的數字）']
        for label in YOY_ROWS:
            cells = []
            for s in [''] + list(DF.STORES[region].values()):
                a, b = V.get((s, f'{label} 本月累計')), V.get((s, f'{label} 去年同期'))
                if a is None or b is None:
                    continue
                cells.append(f'{s or region} {a:,.0f}／{b:,.0f}（{a - b:+,.0f}）')
            L.append(f'【{label}】本月累計／去年同期（差額）：' + '；'.join(cells))
        out[region] = (facts, '\n'.join(L))
    return out


def sent_d43(day):
    """日報附件/日期/寄出_北X區_北X區日報_日期.xlsx 的 D43 概況文字。回傳 {'北一區': text, '北二區': text}。"""
    out = {}
    for p in (ARCHIVE_DAILY / day).glob('寄出_北*區日報_*.xlsx') if (ARCHIVE_DAILY / day).exists() else []:
        region = '北一區' if '北一' in p.name else '北二區'
        try:
            wb = openpyxl.load_workbook(p, read_only=True)
            ws = wb['日報'] if '日報' in wb.sheetnames else wb.worksheets[0]
            t = [c.value.strip() for r in ws.iter_rows() for c in r if isinstance(c.value, str) and len(c.value.strip()) > 30]
            if t:
                out[region] = '\n'.join(t)
        except Exception as e:
            log('讀寄出日報失敗', p.name, e)
    return out


def recent_sent(day, n=7):
    d = dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))
    rows = []
    for i in range(n, 0, -1):
        k = (d - dt.timedelta(days=i)).strftime('%Y%m%d')
        for region, text in sorted(sent_d43(k).items()):
            rows.append((k, region, text))
    return rows


def refresh_corpus(day):
    """把語料最後一天之後、到 day 前一天為止的實寄概況補進 Bob日報.txt／Yuni日報.txt（語氣範例才不會停在舊的）。"""
    end = dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))
    for region, name in (('北一區', 'Bob日報.txt'), ('北二區', 'Yuni日報.txt')):
        p = CORPUS / name
        if not p.exists():
            continue
        text = p.read_text(encoding='utf-8')
        last = max(re.findall(r'(?m)^## (\d{8})', text), default=None)
        if not last:
            continue
        d, add = dt.date(int(last[:4]), int(last[4:6]), int(last[6:])) + dt.timedelta(days=1), []
        while d < end:
            sent = sent_d43(d.strftime('%Y%m%d')).get(region)
            if sent:
                add.append(f'## {d:%Y%m%d}（{"一二三四五六日"[d.weekday()]}）\n{sent}\n')
            d += dt.timedelta(days=1)
        if add:
            p.write_text(text.rstrip('\n') + '\n\n' + '\n'.join(add), encoding='utf-8')
            log('語料補', name, len(add), '天')


def corpus_tail(name, n):
    p = CORPUS / name
    if not p.exists():
        return ''
    parts = re.split(r'(?m)^## ', p.read_text(encoding='utf-8'))
    return '\n'.join('## ' + x.strip() for x in parts[-n:] if x.strip())


def yesterday_compare(day):
    """前一天的自動選項 vs 實寄版（最像哪一則、相似度）；順便補記近 14 天漏記的命中紀錄與改寫紀錄。"""
    d = dt.date(int(day[:4]), int(day[4:6]), int(day[6:]))
    prev = (d - dt.timedelta(days=1)).strftime('%Y%m%d')
    for i in range(14, 1, -1):                      # 補記：實寄晚到、或那天沒跑對比
        k = (d - dt.timedelta(days=i)).strftime('%Y%m%d')
        try:
            sent = sent_d43(k)
            if sent:
                BX.compare_day(k, sent)
        except Exception as e:
            log('補記命中失敗', k, e)
    res = OUT_BASE / '自動盲測' / prev / 'result.json'
    sent = sent_d43(prev)
    if not res.exists() or not sent:
        return f'{prev[4:6]}/{prev[6:]}：' + ('沒有前一天的自動選項' if not res.exists() else '實寄版還沒收到（寄出信沒有附件或還沒寄）')
    shown = BX.compare_day(prev, sent) or ['沒有前一天的選項可比']
    return f'{prev[4:6]}/{prev[6:]} 對比：' + '｜'.join(shown)


# ---------------------------------------------------------------- Claude
def ask_claude(prompt, outdir, tag, retry=True):
    env = {k: v for k, v in os.environ.items() if k != 'ANTHROPIC_API_KEY'}
    if TOKEN.exists():
        env['CLAUDE_CODE_OAUTH_TOKEN'] = TOKEN.read_text().strip()
    (outdir / f'prompt_{tag}.txt').write_text(prompt, encoding='utf-8')
    r = subprocess.run([str(CLAUDE), '-p', '--model', 'opus'], input=prompt, capture_output=True,
                       text=True, timeout=900, env=env)
    (outdir / f'claude_{tag}.txt').write_text(r.stdout + '\n---stderr---\n' + r.stderr, encoding='utf-8')
    if r.returncode:
        raise RuntimeError(f'claude -p 失敗：{r.stderr.strip()[-300:]}')
    m = re.search(r'\{.*\}', r.stdout, re.S)
    try:
        return json.loads(m.group(0))
    except (AttributeError, json.JSONDecodeError):
        if retry:
            return ask_claude(prompt + '\n\n（上次回覆不是合法 JSON，這次只輸出 JSON，字串裡的雙引號要跳脫）',
                              outdir, tag + '_retry', retry=False)
        raise RuntimeError('Claude 回覆裡找不到合法 JSON')


def build_prompt(day, regions, tables, recent, compare):
    md = f'{day[:4]}/{day[4:6]}/{day[6:]}（{"一二三四五六日"[dt.date(int(day[:4]), int(day[4:6]), int(day[6:])).weekday()]}）'
    rules = (HERE / '口徑規格' / '自動盲測規則.md').read_text(encoding='utf-8')
    examples = (CORPUS / 'daily-report-examples.md').read_text(encoding='utf-8') if (CORPUS / 'daily-report-examples.md').exists() else ''
    rec = '\n\n'.join(f'### {k} {r}\n{t}' for k, r, t in recent) or '（沒有資料）'
    parts = [
        f'你要替 STUDIO A 兩位區主管寫 {md} 的日報「本日營運概況」盲測草稿：北一區用 Bob 的語氣、北二區用 Yuni 的語氣。',
        '只輸出一個 JSON，不要其他文字：{"北一區":[{"id":"A","topic":"主題 6～12 字","text":"整段概況"}, ...], "北二區":[...]}。',
        f'每區 {"、".join(regions)} 給 2～3 個選項（A 是你最推薦的），同區選項主題不同；兩區的 A 主題一定不同，B、C 也盡量錯開。',
        '# 規則（一定要遵守）\n' + rules,
        '# 寫法範例與兩人差異\n' + examples,
        '# Bob 近期實寄（語氣參考）\n' + corpus_tail('Bob日報.txt', 18),
        '# Yuni 近期實寄（語氣參考）\n' + corpus_tail('Yuni日報.txt', 18),
        '# 近 7 天兩區實寄概況（選題不能撞自己區這些主題）\n' + rec,
        '# 昨天自動選項 vs 實寄\n' + compare,
        '# Bob／Yuni 怎麼改我的稿（被採用選項 vs 實寄版，照這些改寫方向寫；「換了主題」表示他們寫了我沒想到的題）\n'
        + (BX.rewrite_tail() or '（還沒有紀錄）'),
        '# 其他區區主管近期日報（只參考選題與切入點；他們的寫法、題型不一定適用北一北二，數字不能用）\n'
        + (BX.other_regions(day) or '（沒有資料）'),
    ]
    for region in regions:
        parts.append(f'# {region} 數字（只能用這裡出現的數字，或由它們算出的差額／百分點差）\n' + tables[region])
    return '\n\n'.join(parts)


def fix_prompt(region, opt, bad_rows, table):
    return '\n\n'.join([
        f'下面這則 {region} 日報概況草稿核對時有數字對不回來源（❌）。請只修正這些數字：改成表裡正確的數字，'
        '或刪掉那句；其他文字盡量不動、語氣不變。只輸出 JSON：{"text":"修正後整段"}',
        '草稿：\n' + opt['text'],
        '核對 ❌：\n' + '\n'.join(bad_rows),
        f'{region} 數字表：\n' + table])


def overlap_prompt(opt, dup):
    return '\n\n'.join([
        f'下面這則 {opt["region"]} 日報概況草稿，有片語跟另一區主管最近寫的（或今天另一區的草稿）幾乎一樣，'
        '兩區主管的日報會被同一批長官看到，不能像抄的。請把這些片語換個說法或換切入點，數字一個都不能改、不能增減，'
        '語氣維持原本那位主管。只輸出 JSON：{"text":"改寫後整段"}',
        '草稿：\n' + opt['text'],
        '撞到的片語：\n' + '\n'.join(f'- {x}' for x in dup[:8])])


def check_text(text, facts, region):
    md, bad, warn, _, _ = fc.report(text, facts, PROFILE[region],
                                    fc.N1_STORES if region == '北一區' else fc.N2_STORES, rule_text=text)
    bad_rows = [l for l in md.splitlines() if '❌' in l and not l.startswith(('#', '數字', '**'))]
    return md, bad, warn, bad_rows


# ---------------------------------------------------------------- Discord
def post(content):
    url = WEBHOOK.read_text().strip()
    for i in range(0, len(content), 1900):
        body = json.dumps({'content': content[i:i + 1900]}, ensure_ascii=False).encode()
        req = urllib.request.Request(url, data=body, method='POST', headers={
            'Content-Type': 'application/json', 'User-Agent': 'studioa-daily-blind/1.0'})
        with urllib.request.urlopen(req, timeout=60):
            pass
        time.sleep(1)


# ---------------------------------------------------------------- sa-north-hub
def push_hub(result):
    """把當天的 result.json 推到網站 /daily-blind（管理員頁）；token 跟其他推雲端腳本共用。"""
    token = os.environ.get('EPB_PUSH_TOKEN') or HUB_TOKEN.read_text().strip()
    req = urllib.request.Request(HUB + '/api/daily-blind', data=json.dumps(result, ensure_ascii=False).encode(),
                                 headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())


# ---------------------------------------------------------------- main
def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    dry, wait = '--dry' in sys.argv, '--no-wait' not in sys.argv
    day = args[0] if args else (dt.date.today() - dt.timedelta(days=1)).strftime('%Y%m%d')
    outdir = OUT_BASE / '自動盲測' / day
    if '--push-only' in sys.argv:                 # 不重寫，只把已經產好的那天補推到網站
        result = json.loads((outdir / 'result.json').read_text(encoding='utf-8'))
        result['hits'] = BX.hit_records(day)
        print(push_hub(result))
        return
    outdir.mkdir(parents=True, exist_ok=True)
    md = f'{day[4:6]}/{day[6:]}'
    notes = []
    try:
        miss = wait_for_files(day, wait, notify=None if dry else post)
        check_lines = check_message(day, miss, False).splitlines()[1:]      # 網站跟 Discord 同一份（去掉標題）
        regions = list(DF.STORES)
        for r in DF.STORES:
            if miss[r]:
                notes.append(f'⚠️ {r} 門市日報缺 {"、".join(miss[r])}：照寫，略過缺的店；日報區合計不用（缺店偏低），區數字用 EPB 口徑')

        facts, tables = {}, {}
        for r in regions:
            f, checks = DF.build(day, r, allow_missing=True)
            bad = [c for c in checks if c[0] == '❌']
            if bad:
                notes.append(f'⚠️ {r} 事實檔有 ❌：{bad}，這區今天不寫')
                continue
            facts[r], tables[r] = f, daily_table(f, r)
            if miss[r]:
                tables[r] = (f'⚠️ 今天 {"、".join(miss[r])} 沒有門市日報：不要提這店、不要寫日報口徑的「{r}」合計'
                             f'（下表沒有區合計）；區整體數字只能用後面保固圖／重點項目（EPB 口徑，含全部門市）。\n\n'
                             + tables[r])
        regions = list(facts)
        stores, store_lines = [], []
        try:
            stores = BX.store_verdicts(day, facts)
            store_lines = BX.verdict_lines(stores)
            if not dry:
                post(f'🏪 **{md} 門市概況能不能寫進兩區日報**\n' + '\n'.join(store_lines))
        except Exception as e:
            log('逐店判斷失敗：', e)
        try:
            ac = acpp_facts(day)
            for r in regions:
                facts[r] += ac[r][0]
                tables[r] += '\n\n（保固圖口徑：只算 ACPP+，排除 SA Care 檢測新機與認證機；跟每天 12:00 推的保固圖一樣）\n' + ac[r][1]
        except Exception as e:
            notes.append(f'⚠️ 保固圖口徑查不到（{type(e).__name__}: {e}），今天只用日報口徑')
        extra = []
        try:
            extra.append(BX.store_notes(day))
        except Exception as e:
            notes.append(f'⚠️ 門市概況讀不到（{type(e).__name__}: {e}）')
        try:
            extra.append(BX.bulk_orders(day, facts, _epb_rows, _base_cond, _shop))
        except Exception as e:
            notes.append(f'⚠️ 大單查不到（{type(e).__name__}: {e}）')
        if day.endswith('01'):
            try:
                extra.append(BX.month_facts(day))
            except Exception as e:
                notes.append(f'⚠️ 上月整月數字查不到（{type(e).__name__}: {e}），今天沒有整月概況')
        for ex in extra:
            for r in regions:
                if r in ex and ex[r][1]:
                    facts[r] += ex[r][0]
                    tables[r] += '\n\n' + ex[r][1]
        try:
            fo = focus_facts(day)
            for r in regions:
                if fo[r][1]:
                    facts[r] += fo[r][0]
                    tables[r] += '\n\n' + fo[r][1]
        except Exception as e:
            notes.append(f'⚠️ 月重點項目／新品查不到（{type(e).__name__}: {e}）')
        try:                                   # SA Care 價目表跟 Google 表同步（週對週引擎用）
            r = subprocess.run([sys.executable, str(SCRIPTS / 'sacare_price_sync.py')], capture_output=True,
                               text=True, timeout=180)
            log('價目表同步', r.stdout.strip()[-300:], r.stderr.strip()[-300:])
        except Exception as e:
            log('價目表同步失敗', e)
        if dt.date(int(day[:4]), int(day[4:6]), int(day[6:])).weekday() == 5:
            try:
                wk = week_facts(day)
                for r in regions:
                    facts[r] += wk[r][0]
                    tables[r] += '\n\n# 週六：本週 vs 上週（A 一定寫週對週）\n' + wk[r][1]
            except Exception as e:
                notes.append(f'⚠️ 週對週數字查不到（{type(e).__name__}: {e}），今天沒有週對週選項')
        if not day.endswith('01'):
            try:
                yy = yoy_facts(day)
                for r in regions:
                    facts[r] += yy[r][0]
                    tables[r] += '\n\n# 本月累計 vs 去年同期\n' + yy[r][1]
            except Exception as e:
                notes.append(f'⚠️ 去年同期數字查不到（{type(e).__name__}: {e}），今天不寫去年同期')
        for r in regions:
            fc.save_facts(str(outdir / f'facts_{r}.json'), facts[r], {'kind': 'daily_auto', 'region': r, 'day': day})

        try:                                   # 其他區主管日報 → 語料（mini 從信件備份補新的）
            subprocess.run([sys.executable, str(HERE / '其他區日報_語料.py')], capture_output=True, text=True, timeout=300)
        except Exception as e:
            log('其他區語料更新失敗', e)
        compare = yesterday_compare(day)
        try:
            refresh_corpus(day)
        except Exception as e:
            log('語料更新失敗', e)
        prompt = build_prompt(day, regions, tables, recent_sent(day), compare)
        ans = ask_claude(prompt, outdir, 'draft')

        options = []
        for r in regions:
            for o in ans.get(r, [])[:3]:
                o = {'region': r, 'id': o.get('id', '?'), 'topic': o.get('topic', ''), 'text': o.get('text', '').strip()}
                md_, bad, warn, bad_rows = check_text(o['text'], facts[r], r)
                if bad:
                    try:
                        fixed = ask_claude(fix_prompt(r, o, bad_rows, tables[r]), outdir, f'fix_{r}_{o["id"]}')
                        o['text_before_fix'] = o['text']
                        o['text'] = fixed['text'].strip()
                        md_, bad, warn, bad_rows = check_text(o['text'], facts[r], r)
                    except Exception as e:
                        log('修正失敗', r, o['id'], e)
                n_nums = len(re.findall(r'^\| (✅|⚠️|❌|🎯) ', md_, re.M))
                o.update(ok=bad == 0, bad=bad, warn=warn, nums=n_nums)
                (outdir / f'{r}_{o["id"]}.txt').write_text(o['text'], encoding='utf-8')
                (outdir / f'核對單_{r}_{o["id"]}.md').write_text(md_, encoding='utf-8')
                options.append(o)
        recent = recent_sent(day)
        for o in options:
            if not o['ok']:
                continue
            others = [t for k, r, t in recent if r != o['region']] + \
                     [x['text'] for x in options if x['region'] != o['region']]
            dup = BX.overlaps(o['text'], others)
            if dup:
                try:
                    fixed = ask_claude(overlap_prompt(o, dup), outdir, f'dedup_{o["region"]}_{o["id"]}')
                    md_, bad, warn, _ = check_text(fixed['text'].strip(), facts[o['region']], o['region'])
                    if bad == 0:
                        o['text_before_dedup'], o['text'] = o['text'], fixed['text'].strip()
                        o.update(warn=warn)
                        (outdir / f'{o["region"]}_{o["id"]}.txt').write_text(o['text'], encoding='utf-8')
                        (outdir / f'核對單_{o["region"]}_{o["id"]}.md').write_text(md_, encoding='utf-8')
                        dup = BX.overlaps(o['text'], others)
                except Exception as e:
                    log('撞句改寫失敗', o['region'], o['id'], e)
            o['dup'] = dup
        hit = BX.hit_summary(day)
        result = {'day': day, 'options': options, 'notes': notes, 'compare': compare, 'hit': hit,
                  'check': check_lines, 'stores': stores, 'store_lines': store_lines,
                  'hits': BX.hit_records(day)}
        (outdir / 'result.json').write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding='utf-8')
        if not dry:
            try:
                log('推網站', push_hub(result))
            except Exception as e:
                log('推網站失敗（Discord 照推）：', e)

        head = [f'📝 **{md} 兩區日報概況盲測**（核對 0 ❌ 才列出；草稿與核對單在 mini `輸出/核對/自動盲測/{day}/`）', compare,
                hit] + notes
        head = [h for h in head if h]
        msgs = ['\n'.join(head)]
        for o in options:
            if not o['ok']:
                msgs.append(f'~~{o["region"][:2]} {o["id"]}｜{o["topic"]}~~ 核對仍有 ❌ {o["bad"]} 個，已剔除')
                continue
            w = f'⚠️ {o["warn"]}（看核對單）' if o['warn'] else '全部通過'
            if o.get('dup'):
                w += f'｜跟另一區撞句：「{o["dup"][0][:30]}」'
            msgs.append(f'**{o["region"][:2]} {o["id"]}｜{o["topic"]}**　數字 {o["nums"]} 個 0 ❌｜{w}\n```\n{o["text"]}\n```')
        text = '\n\n'.join(msgs)
        print(text)
        if not dry:
            for m in msgs:
                post(m)
    except Exception as e:
        msg = f'⚠️ {md} 日報概況盲測沒產出：{type(e).__name__}: {e}'
        log(msg)
        if not dry:
            try:
                post(msg)
            except Exception as e2:
                log('失敗通知也推不出去：', e2)
        sys.exit(1)


if __name__ == '__main__':
    main()
