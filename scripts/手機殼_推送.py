#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
北區 iPhone 18 手機殼 — 查 EPB 推到 sa-case-hub
==============================================
EPB 只走公司 VPN，雲端打不到：本機查完庫存、日銷、在途後 POST 上去。
launchd 每天 11:00、22:00 跑；VPN 沒連就記 log 跳過。

【用法】
    python3 手機殼_推送.py                         # 查 EPB 推上去
    python3 手機殼_推送.py --dry-run               # 只查不推，印出筆數
    python3 手機殼_推送.py --seed-xlsx 檔案.xlsx --dry-run   # 印出配貨單的掛勾比對（不查 EPB）
    python3 手機殼_推送.py --seed-xlsx 檔案.xlsx   # 把配貨單推上去
    python3 手機殼_推送.py --poll                  # 網站有人按「立即更新」才查（launchd 每分鐘跑）
    HUB_URL=http://127.0.0.1:5060 python3 手機殼_推送.py      # 推本機測試站

【Token】環境變數 CASE_PUSH_TOKEN，或 ~/.config/studioa/case_push_token
==============================================
"""
import argparse
import csv
import io
import json
import os
import ssl
import subprocess
import sys
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta

HUB_REPO = os.path.expanduser('~/sa_case_hub')
sys.path.insert(0, HUB_REPO)
import case_styles as cs  # noqa: E402
import db  # noqa: E402
from stores import STORE_CODES, STORE_NAME  # noqa: E402

HUB_URL = os.environ.get('HUB_URL', 'https://sa-case-hub.fly.dev').rstrip('/')
TOKEN_FILE = os.path.expanduser('~/.config/studioa/case_push_token')
JAVA = '/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java'
CP = ':'.join([os.path.expanduser('~/工具中心/lib'), '/Library/EPBrowser/EPB/Shell/shell.jar',
               '/Library/EPBrowser/EPB/Shell/lib/*'])
EPB_ROOT = '/Library/EPBrowser'
SALES_FLOOR = date(2026, 9, 1)          # 日銷最早只推到這天，之前沒有 18 殼
CASE_CAT6 = "('6053','6054','6055','6056')"


def token():
    value = os.environ.get('CASE_PUSH_TOKEN')
    if not value and os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, encoding='utf-8') as handle:
            value = handle.read().strip()
    if not value:
        sys.exit(f'找不到 token：設環境變數 CASE_PUSH_TOKEN 或寫進 {TOKEN_FILE}')
    return value


def epb(sql, max_rows=200000):
    """跑一段唯讀 SQL，回傳 list[dict]。超過 max_rows 直接中止，不要靜默截斷。"""
    proc = subprocess.run([JAVA, '-Dsun.net.client.defaultConnectTimeout=8000',
                           '-Dsun.net.client.defaultReadTimeout=300000',
                           '-cp', CP, 'EPBReportQuery', EPB_ROOT, sql, str(max_rows)],
                          capture_output=True, text=True, timeout=600)
    if 'MAX_ROWS_REACHED' in proc.stderr:
        raise RuntimeError(f'EPB 回傳筆數達上限 {max_rows}，資料會被截斷')
    if proc.returncode != 0 or not proc.stdout.strip():
        if proc.returncode != 0:
            raise RuntimeError(f'EPB 查詢失敗（VPN 有連嗎？）：{proc.stderr.strip()[-300:]}')
        return []
    reader = csv.DictReader(io.StringIO(proc.stdout), delimiter='\t', quoting=csv.QUOTE_NONE)
    return [{k.upper(): (v or '').strip() for k, v in row.items()} for row in reader]


def in_list(values):
    return '(' + ','.join("'" + str(v).replace("'", "''") + "'" for v in values) + ')'


def http(path, payload=None):
    data = json.dumps(payload).encode('utf-8') if payload is not None else None
    request = urllib.request.Request(f'{HUB_URL}{path}', data=data, method='POST' if data else 'GET',
                                     headers={'Content-Type': 'application/json',
                                              'Authorization': f'Bearer {token()}'})
    # python.org 版 Python 沒帶系統根憑證，urllib 連 https 會驗證失敗；有 certifi 就用它的
    try:
        import certifi
        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        context = None
    try:
        with urllib.request.urlopen(request, timeout=120, context=context) as response:
            return json.load(response)
    except urllib.error.HTTPError as error:
        sys.exit(f'❌ {path} HTTP {error.code}：{error.read().decode("utf-8", "replace")}')
    except urllib.error.URLError as error:
        sys.exit(f'❌ 連不到 {HUB_URL}：{error.reason}')


def fetch(today, gift_ids, settings_days=10):
    locs = in_list(STORE_CODES)
    stores = in_list('SA' + c for c in STORE_CODES)
    case_filter = f"m.cat5_id IN ('5609','5610') AND m.cat1_id IN ('1001','1002') AND m.cat6_id IN {CASE_CAT6}"
    gift_filter = f' OR m.stk_id IN {in_list(gift_ids)}' if gift_ids else ''

    sku_rows = epb(f"SELECT m.stk_id, m.name, m.cat5_id, m.cat1_id, m.brand_id FROM stkmas m WHERE ({case_filter}{gift_filter})")
    skus, gifts = [], []
    for r in sku_rows:
        if r['STK_ID'] in gift_ids:
            gifts.append({'stk_id': r['STK_ID'], 'name': r['NAME'].lstrip('@')})
        model = cs.CAT5_MODEL.get(r['CAT5_ID'])
        if model and r['CAT1_ID'] in ('1001', '1002'):
            # 廠商名 EPB 主檔沒有（只有品牌代碼），留空白讓雲端保留配貨單帶進來的廠商名
            skus.append({'stk_id': r['STK_ID'], 'name': r['NAME'].lstrip('@'), 'vendor': '',
                         'model': model, 'kind': 'demo' if r['CAT1_ID'] == '1002' else 'normal'})
    stk_ids = [s['stk_id'] for s in skus] + list(gift_ids)
    if not stk_ids:
        return skus, gifts, [], [], [], SALES_FLOOR

    stock = [{'store_id': r['STORE_ID'][2:], 'stk_id': r['STK_ID'], 'qty': float(r['Q'])}
             for r in epb(f"SELECT s.store_id, s.stk_id, SUM(s.stk_qty) q FROM storesum s, stkmas m"
                          f" WHERE s.stk_id=m.stk_id AND s.store_id IN {stores} AND ({case_filter}{gift_filter})"
                          f" GROUP BY s.store_id, s.stk_id HAVING SUM(s.stk_qty)<>0")]

    sales_from = max(SALES_FLOOR, today - timedelta(days=35))
    sales = [{'sale_date': r['D'], 'store_id': r['SHOP_ID'].zfill(3), 'stk_id': r['STK_ID'],
              'qty': float(r['Q']), 'gift_qty': float(r['G'])}
             for r in epb(f"SELECT TO_CHAR(l.doc_date,'YYYY-MM-DD') d, l.shop_id, l.stk_id, SUM(l.stk_qty) q,"
                          f" SUM(CASE WHEN l.trans_type='A' AND l.line_total_net=0 THEN l.stk_qty ELSE 0 END) g"
                          f" FROM poslinev_bi l, stkmas m WHERE l.stk_id=m.stk_id AND l.org_id='01'"
                          f" AND l.shop_id IN {locs} AND l.trans_type IN ('A','E','H')"
                          f" AND l.doc_date>=TO_DATE('{sales_from}','YYYY-MM-DD') AND ({case_filter}{gift_filter})"
                          f" GROUP BY TO_CHAR(l.doc_date,'YYYY-MM-DD'), l.shop_id, l.stk_id")]

    since = today - timedelta(days=settings_days)
    pr = [(r['LOC_ID'], r['STK_ID'], datetime.strptime(r['D'], '%Y-%m-%d').date(), float(r['Q']))
          for r in epb(f"SELECT p.loc_id, l.stk_id, TO_CHAR(p.doc_date,'YYYY-MM-DD') d, SUM(l.stk_qty) q"
                       f" FROM prmas p, prline l, stkmas m WHERE p.rec_key=l.mas_rec_key AND l.stk_id=m.stk_id"
                       f" AND p.loc_id IN {locs} AND p.doc_date>=TO_DATE('{since}','YYYY-MM-DD') AND ({case_filter})"
                       f" GROUP BY p.loc_id, l.stk_id, TO_CHAR(p.doc_date,'YYYY-MM-DD')")]
    inbound = [(r['STORE_ID'][2:], r['STK_ID'], datetime.strptime(r['D'], '%Y-%m-%d').date(), float(r['Q']))
               for r in epb(f"SELECT d.store_id, d.stk_id, TO_CHAR(d.doc_date,'YYYY-MM-DD') d, SUM(d.stk_qty) q"
                            f" FROM storedtl d, stkmas m WHERE d.stk_id=m.stk_id AND d.store_id IN {stores}"
                            f" AND d.src_code IN ('GRN','INVTRNTN') AND d.stk_qty>0"
                            f" AND d.doc_date>=TO_DATE('{since}','YYYY-MM-DD') AND ({case_filter})"
                            f" GROUP BY d.store_id, d.stk_id, TO_CHAR(d.doc_date,'YYYY-MM-DD')")]
    transit = [{'store_id': s, 'stk_id': k, 'requested': a, 'received': b, 'in_transit': c}
               for (s, k), (a, b, c) in cs.compute_in_transit(pr, inbound, today, settings_days).items()]
    # 主檔的 18 殼有好幾千個 SKU，大多沒配到北區；只推北區真的有庫存/銷售/在途的，免得款式主檔塞滿
    relevant = {r['stk_id'] for r in stock} | {r['stk_id'] for r in sales} | {r['stk_id'] for r in transit}
    skus = [s for s in skus if s['stk_id'] in relevant]
    return skus, gifts, stock, sales, transit, sales_from


def seed_report(path):
    conn = db.connect(':memory:')
    db.init_db(conn)
    with open(path, 'rb') as handle:
        hooks, alloc, skus = cs.parse_allocation_workbook(handle, STORE_CODES)
    cs.load_allocation(conn, hooks, alloc, skus, 'seed')
    print(f'配貨單：{len(alloc)} 筆配貨、{len(skus)} 個 SKU\n')
    print(f'{"門市":<6}{"機型":<8}{"掛勾":>4}{"demo件":>7}{"差":>4}{"demo款":>7}{"正常品款":>8}  異常')
    for code in STORE_CODES:
        comp = cs.compliance(conn, code, date(2026, 9, 14))
        for model in cs.MODELS:
            c = comp[model]
            kinds = {}
            for e in c['exceptions']:
                kinds[e['label']] = kinds.get(e['label'], 0) + 1
            print(f'{STORE_NAME[code]:<6}{model:<8}{c["hooks"]:>4}{c["demo_alloc"]:>7g}{c["demo_alloc"] - c["hooks"]:>+4g}'
                  f'{c["demo_styles"]:>7}{c["normal_styles"]:>8}  {kinds}')
    return hooks, alloc, skus


def push_epb(stamp, transit_days, dry_run=False):
    """查 EPB 並推上去。回傳 (成功與否, 說明)，給排程與「立即更新」共用。"""
    gift_ids = [] if dry_run and not os.path.exists(TOKEN_FILE) and not os.environ.get('CASE_PUSH_TOKEN') \
        else http('/api/gift-skus').get('stk_ids', [])
    try:
        skus, gifts, stock, sales, transit, sales_from = fetch(stamp.date(), gift_ids, transit_days)
    except (RuntimeError, subprocess.TimeoutExpired) as error:
        message = f'EPB 查不到（VPN 有連嗎？）：{str(error)[:150]}'
        print(f'⚠️ 跳過這次推送：{message}')
        return False, message
    summary = (f'SKU {len(skus)}（demo {sum(1 for s in skus if s["kind"] == "demo")}）、庫存 {len(stock)} 筆、'
               f'日銷 {len(sales)} 筆（{sales_from} 起）、在途 {len(transit)} 筆、贈品 {len(gifts)}')
    print(summary)
    if not dry_run:
        print(http('/api/push', {'snap_at': f'{stamp:%Y-%m-%d %H:%M:%S}', 'sales_from': sales_from.isoformat(),
                                 'skus': skus, 'gifts': gifts, 'stock': stock, 'sales_day': sales,
                                 'in_transit': transit, 'pushed_by': os.environ.get('USER', '')}))
    return True, summary


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--seed-xlsx')
    parser.add_argument('--dry-run', action='store_true')
    parser.add_argument('--poll', action='store_true', help='看網站有沒有人按「立即更新」，有才查 EPB')
    parser.add_argument('--transit-days', type=int, default=10)
    args = parser.parse_args()
    stamp = datetime.now(cs.TZ)

    if args.poll:
        request_id = http('/api/sync-request').get('id')
        if not request_id:
            return                      # 每分鐘跑一次，沒人按就安靜結束，不要洗 log
        print(f'[{stamp:%Y-%m-%d %H:%M:%S}] 網站按了立即更新（#{request_id}）→ {HUB_URL}')
        try:
            ok, message = push_epb(stamp, args.transit_days)
        except SystemExit as error:     # http() 連線失敗會 sys.exit，要回報失敗而不是讓網站一直轉圈
            ok, message = False, str(error)
        http('/api/sync-result', {'id': request_id, 'ok': ok, 'message': message})
        return

    print(f'[{stamp:%Y-%m-%d %H:%M:%S}] 手機殼推送 → {HUB_URL}')
    if args.seed_xlsx:
        hooks, alloc, skus = seed_report(args.seed_xlsx)
        if not args.dry_run:
            print(http('/api/allocation', {'hooks': hooks, 'alloc': alloc, 'skus': skus,
                                           'filename': os.path.basename(args.seed_xlsx),
                                           'pushed_by': os.environ.get('USER', '')}))
        return
    push_epb(stamp, args.transit_days, args.dry_run)


if __name__ == '__main__':
    main()
