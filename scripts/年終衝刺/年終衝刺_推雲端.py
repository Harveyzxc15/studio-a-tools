#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
年終衝刺 — 本機查 EPB，全台 33 店的總業績／3PP／保固推到 sa-north-hub（/sprint）
  口徑＝主管 studioa-dashboard「Lifestyle 銷售差異」（2026-11-15 逐月對過：5 區 × 23 個月，3PP、Apple Care 逐元相同，
        總業績只有 2026/9 儀表板少 3–5%（它的快照比較早），保固差 0.06% 以內）：
    金額＝poslinev_bi line_total_net＋line_tax，trans_type A/E/G/H/J，org 01
    總業績＝全部金額＋SA Care 估算；3PP＝cat3 3003；Apple Care＝cat3 3032
    SA Care 估算＝cat3 3012 品名「($2,090)」括號定價 × stk_qty，只算 A/E/H（成交金額欄是 0）
    保固＝SA Care 估算＋Apple Care
  推：buckets（Y26＝今年 1/1～11/14、Y25＝去年同期、S26/S25＝9/1～11/14、N26/N25＝11/1～11/14）＋
      daily（去年 11/15～12/31 與今年 11/15～昨天，每天每店）
  年度成長、名次、門市目標都在網站 sprint.py 算

用法：python3 年終衝刺_推雲端.py            # 直接推
      python3 年終衝刺_推雲端.py --poll     # 有人按「刷新」，或今天 08:00 後還沒推過才推（launchd 每 2 分鐘跑）
      python3 年終衝刺_推雲端.py --dry-run  # 只查不推，印各區合計
Token：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token（與其他推送共用）
"""
import csv, io, json, os, re, ssl, subprocess, sys, urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta

HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
JAVA = '/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java'
LIB = os.path.expanduser('~/工具中心/lib')
EPB_CP = f'{LIB}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*'
SHOPS = [3, 4, 5, 9, 12, 21, 24, 25, 27, 31, 35, 37, 43, 44, 46, 47, 48, 50, 51, 52, 53, 54, 55, 57,
         58, 61, 62, 63, 64, 65, 66, 67, 68]
DAILY_AFTER = '08:00'
BUCKETS = {  # 名稱: (起, 迄) 含頭含尾
    'Y26': ('2026-01-01', '2026-11-14'), 'Y25': ('2025-01-01', '2025-11-14'),
    'S26': ('2026-09-01', '2026-11-14'), 'S25': ('2025-09-01', '2025-11-14'),
    'N26': ('2026-11-01', '2026-11-14'), 'N25': ('2025-11-01', '2025-11-14'),
}
LY_DAYS = ('2025-11-15', '2025-12-31')
TY_START, TY_END = date(2026, 11, 15), date(2026, 12, 31)
PRICE = re.compile(r'\(\$([\d,]+)\)')


def token():
    t = os.environ.get('EPB_PUSH_TOKEN')
    if not t:
        t = open(os.path.expanduser('~/.config/studioa/epb_push_token')).read().strip()
    return t


def call(path, body=None):
    req = urllib.request.Request(
        HUB + path, data=None if body is None else json.dumps(body, ensure_ascii=False).encode(),
        headers={'Authorization': 'Bearer ' + token(), 'Content-Type': 'application/json'})
    try:
        import certifi
        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        context = None
    with urllib.request.urlopen(req, timeout=60, context=context) as r:
        return json.loads(r.read())


def run_sql(sql, max_rows=300000):
    r = subprocess.run([JAVA, '-cp', EPB_CP, 'EPBReportQuery', '/Library/EPBrowser', sql, str(max_rows)],
                       capture_output=True, text=True, cwd=LIB, timeout=900)
    if 'MAX_ROWS_REACHED' in r.stderr:
        raise SystemExit(f'EPB 回傳筆數達上限 {max_rows}，資料會被截斷')
    if r.returncode != 0:
        raise SystemExit('EPB 查詢失敗（VPN 沒連？）：' + r.stderr.strip()[-300:])
    reader = csv.DictReader(io.StringIO(r.stdout), delimiter='\t', quoting=csv.QUOTE_NONE)
    return [{k.upper(): (v or '').strip() for k, v in row.items()} for row in reader]


def d(s):
    return f"TO_DATE('{s}','YYYY-MM-DD')"


def fetch(yesterday):
    """回傳 (buckets, daily)。每個時段用 CASE 標成一個鍵：bucket 名稱，或 'D' 加日期（逐日）。"""
    ty_end = min(yesterday, TY_END)
    cases = [f"WHEN doc_date BETWEEN {d(a)} AND {d(b)} THEN '{n}'" for n, (a, b) in BUCKETS.items()
             if n not in ('S26', 'S25', 'N26', 'N25')]
    day_cond = f"doc_date BETWEEN {d(LY_DAYS[0])} AND {d(LY_DAYS[1])}"
    if ty_end >= TY_START:
        day_cond = f"({day_cond} OR doc_date BETWEEN {d(TY_START)} AND {d(ty_end)})"
    # Y 時段跟 S/N 時段重疊，S/N 另外從逐日以外的月份算不到，所以 S/N 用第二支查詢
    scope = (f"org_id='01' AND shop_id IN ({','.join(map(str, SHOPS))}) AND trans_type IN ('A','E','G','H','J')")
    amt = 'line_total_net+line_tax'

    def query(key_expr, where):
        base = run_sql(
            f"SELECT {key_expr} k, shop_id, SUM({amt}) t, SUM(CASE WHEN cat3_id='3003' THEN {amt} ELSE 0 END) p,"
            f" SUM(CASE WHEN cat3_id='3032' THEN {amt} ELSE 0 END) ac FROM poslinev_bi WHERE {scope} AND {where}"
            f" GROUP BY {key_expr}, shop_id")
        sa = run_sql(
            f"SELECT {key_expr} k, shop_id, name, SUM(stk_qty) q FROM poslinev_bi WHERE {scope} AND {where}"
            f" AND cat3_id='3012' AND trans_type IN ('A','E','H') GROUP BY {key_expr}, shop_id, name")
        est = defaultdict(float)
        for r in sa:
            m = PRICE.search(r['NAME'])
            if m:
                est[(r['K'], r['SHOP_ID'].zfill(3))] += int(m.group(1).replace(',', '')) * float(r['Q'] or 0)
        out = {}
        for r in base:
            k = (r['K'], r['SHOP_ID'].zfill(3))
            e, ac = est.get(k, 0), float(r['AC'] or 0)
            out[k] = [round(float(r['T'] or 0) + e), round(float(r['P'] or 0)), round(e + ac)]
        return out

    buckets = defaultdict(dict)
    y_key = 'CASE ' + ' '.join(cases) + ' END'
    for (k, s), v in query(y_key, f"doc_date BETWEEN {d('2025-01-01')} AND {d('2026-11-14')}"
                                  f" AND (doc_date<={d('2025-11-14')} OR doc_date>={d('2026-01-01')})").items():
        buckets[k][s] = v
    sn_key = (f"CASE WHEN doc_date BETWEEN {d('2026-11-01')} AND {d('2026-11-14')} THEN 'N26'"
              f" WHEN doc_date BETWEEN {d('2025-11-01')} AND {d('2025-11-14')} THEN 'N25'"
              f" WHEN doc_date BETWEEN {d('2026-09-01')} AND {d('2026-10-31')} THEN 'S26x'"
              f" WHEN doc_date BETWEEN {d('2025-09-01')} AND {d('2025-10-31')} THEN 'S25x' END")
    sn = query(sn_key, f"(doc_date BETWEEN {d('2025-09-01')} AND {d('2025-11-14')}"
                       f" OR doc_date BETWEEN {d('2026-09-01')} AND {d('2026-11-14')})")
    for (k, s), v in sn.items():
        if k in ('N26', 'N25'):
            buckets[k][s] = v
    for yy in ('26', '25'):           # S＝9/1～10/31 ＋ 11/1～11/14
        for s in {s for (k, s) in sn}:
            a, b = sn.get(('S' + yy + 'x', s), [0, 0, 0]), sn.get(('N' + yy, s), [0, 0, 0])
            buckets['S' + yy][s] = [x + y for x, y in zip(a, b)]
    daily = [[k, s, *v] for (k, s), v in sorted(query("TO_CHAR(doc_date,'YYYY-MM-DD')", day_cond).items())]
    return dict(buckets), daily


def epb_up():
    import socket
    try:
        with socket.create_connection(('192.168.1.177', 8080), timeout=1.5):
            return True
    except OSError:
        return False


def main():
    if '--dry-run' not in sys.argv:
        pending = call('/api/sprint/pending')
        if '--poll' in sys.argv:
            now = datetime.now()
            daily_due = pending.get('pushed_on') != str(now.date()) and now.strftime('%H:%M') >= DAILY_AFTER
            if not (pending.get('pending') or daily_due) or not epb_up():
                return
    yesterday = date.today() - timedelta(days=1)
    buckets, daily = fetch(yesterday)
    body = {'asof': str(min(yesterday, TY_END)), 'buckets': buckets, 'daily': daily}
    if '--dry-run' in sys.argv:
        regions = {'北一': ['004', '005', '024', '046', '054', '057'], '北二': ['009', '025', '050', '055', '063', '064', '068'],
                   '桃竹': ['035', '037', '043', '048', '053', '058', '061', '065'], '中區': ['031', '047', '051', '052', '062', '066'],
                   '南區': ['003', '012', '021', '027', '044', '067']}
        for reg, shops in regions.items():
            print(reg, {b: [sum((buckets.get(b, {}).get(s) or [0, 0, 0])[i] for s in shops) for i in range(3)] for b in BUCKETS},
                  '去年11/15-12/31', [sum(r[2 + i] for r in daily if r[1] in shops and r[0] < '2026') for i in range(3)])
        json.dump(body, open(os.path.expanduser('~/工具中心/輸出/年終衝刺/sprint_snapshot.json'), 'w'), ensure_ascii=False)
        return
    print(call('/api/sprint/snapshot', body), f'（資料到 {body["asof"]}，{len(daily)} 筆逐日）')


if __name__ == '__main__':
    main()
