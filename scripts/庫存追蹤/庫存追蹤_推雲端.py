#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
庫存追蹤 — 本機查 EPB，快照推到 sa-north-hub（/stock-watch）
品項在網頁上加（存貨代碼），這支向雲端領全部代碼，查完依代碼推上去：
  庫存＝STOREDTL 全期淨異動（依 store_id=SA+門市碼，同 PQI 已對過 140EB 的口徑）
  銷量＝poslinev_bi 北區 org 01、trans_type A/E/H 的 stk_qty 加總（銷退帶負號；含 0 元贈送），依 shop_id
        本月＝1 號～今天；近 7 天＝今天往回 6 天～今天

用法：python3 庫存追蹤_推雲端.py            # 直接推
      python3 庫存追蹤_推雲端.py --poll     # 雲端有人按「刷新」或加品項才推（launchd 每 2 分鐘跑）
      python3 庫存追蹤_推雲端.py --dry-run  # 只查不推，印出結果
      HUB_URL=http://127.0.0.1:5050 python3 庫存追蹤_推雲端.py      # 推本機測試站
備援機（筆電 launchd 設 SYNC_STANDBY=1）：刷新請求掛滿 3 分鐘、mini 都沒推才接手
Token：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token（與 EPB 實績推送共用）
"""
import csv, io, json, os, ssl, subprocess, sys, urllib.request
from datetime import date, timedelta

HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
JAVA = '/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java'
LIB = os.path.expanduser('~/工具中心/lib')          # EPBReportQuery.class（不放 Desktop：launchd 讀不到）
EPB_CP = f'{LIB}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*'
# 13 店（與 sa_north_hub/stores.py 同一份）；mini 上沒有 git，自含一份免得 import 不到
STORES = ['004', '005', '046', '024', '054', '057', '068', '063', '025', '050', '009', '055', '064']


def token():
    t = os.environ.get('EPB_PUSH_TOKEN')
    if not t:
        t = open(os.path.expanduser('~/.config/studioa/epb_push_token')).read().strip()
    return t


def call(path, body=None):
    req = urllib.request.Request(
        HUB + path, data=None if body is None else json.dumps(body, ensure_ascii=False).encode(),
        headers={'Authorization': 'Bearer ' + token(), 'Content-Type': 'application/json'})
    # python.org 版 Python 沒帶系統根憑證，有 certifi 就用它的
    try:
        import certifi
        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        context = None
    with urllib.request.urlopen(req, timeout=60, context=context) as r:
        return json.loads(r.read())


STANDBY_STATE = os.path.expanduser('~/.config/studioa/stock_watch_standby_seen.json')


def standby_ready(pending):
    """同 PQI兌換_推雲端.py：備援機等請求掛滿 3 分鐘才做；主機（沒設 SYNC_STANDBY）一律直接做。"""
    if os.environ.get('SYNC_STANDBY') != '1':
        return pending
    if not pending:                     # 沒在等了（主機推完）就忘掉，下次重新起算
        if os.path.exists(STANDBY_STATE):
            os.remove(STANDBY_STATE)
        return pending
    import time
    try:
        first = json.load(open(STANDBY_STATE))['first']
    except (OSError, ValueError, KeyError):
        first = time.time()
        json.dump({'first': first}, open(STANDBY_STATE, 'w'))
    return time.time() - first >= 3 * 60


def run_sql(sql, max_rows=50000):
    r = subprocess.run([JAVA, '-cp', EPB_CP, 'EPBReportQuery', '/Library/EPBrowser', sql, str(max_rows)],
                       capture_output=True, text=True, cwd=LIB, timeout=600)
    if 'MAX_ROWS_REACHED' in r.stderr:
        raise SystemExit(f'EPB 回傳筆數達上限 {max_rows}，資料會被截斷')
    if r.returncode != 0:
        raise SystemExit('EPB 查詢失敗（VPN 沒連？）：' + r.stderr.strip()[-300:])
    reader = csv.DictReader(io.StringIO(r.stdout), delimiter='\t', quoting=csv.QUOTE_NONE)
    return [{k.upper(): (v or '').strip() for k, v in row.items()} for row in reader]


def in_list(values):
    return '(' + ','.join("'" + str(v).replace("'", "''") + "'" for v in values) + ')'


def fetch(codes, today):
    d1, d7 = today.replace(day=1), today - timedelta(days=6)
    since = min(d1, d7)
    ids = in_list(codes)
    out = {c: {'name': '', 'stock': {}, 'mtd': {}, 'd7': {}} for c in codes}
    for r in run_sql(f'SELECT stk_id, name FROM stkmas WHERE stk_id IN {ids}'):
        out[r['STK_ID']]['name'] = r['NAME'].lstrip('@')
    for r in run_sql(f'SELECT store_id, stk_id, SUM(stk_qty) q FROM storedtl WHERE stk_id IN {ids}'
                     f" AND store_id IN {in_list('SA' + s for s in STORES)} GROUP BY store_id, stk_id"):
        q = int(float(r['Q'] or 0))
        if q:
            out[r['STK_ID']]['stock'][r['STORE_ID'][2:]] = q
    for r in run_sql(f"SELECT l.shop_id, l.stk_id,"
                     f" SUM(CASE WHEN l.doc_date>=TO_DATE('{d1}','YYYY-MM-DD') THEN l.stk_qty ELSE 0 END) mtd,"
                     f" SUM(CASE WHEN l.doc_date>=TO_DATE('{d7}','YYYY-MM-DD') THEN l.stk_qty ELSE 0 END) d7"
                     f" FROM poslinev_bi l WHERE l.org_id='01' AND l.trans_type IN ('A','E','H')"
                     f" AND l.shop_id IN {in_list(STORES)} AND l.stk_id IN {ids}"
                     f" AND l.doc_date>=TO_DATE('{since}','YYYY-MM-DD')"
                     f" AND l.doc_date<TO_DATE('{today + timedelta(days=1)}','YYYY-MM-DD')"
                     f" GROUP BY l.shop_id, l.stk_id"):
        store = r['SHOP_ID'].zfill(3)              # shop_id 是數字型，回來沒補零
        for key in ('mtd', 'd7'):
            q = int(float(r[key.upper()] or 0))
            if q:
                out[r['STK_ID']][key][store] = q
    missing = [c for c in codes if not out[c]['name']]
    for c in missing:                              # 主檔查無＝代碼打錯，不推空殼
        del out[c]
    return {'range': {'mtd': [str(d1), str(today)], 'd7': [str(d7), str(today)]},
            'codes': out, 'missing': missing}


def main():
    pending = call('/api/stock-watch/pending')
    if '--poll' in sys.argv and not standby_ready(pending.get('pending')):
        return
    codes = pending.get('codes') or []
    body = fetch(codes, date.today()) if codes else {'range': None, 'codes': {}, 'missing': []}
    if '--dry-run' in sys.argv:
        print(json.dumps(body, ensure_ascii=False, indent=1))
        return
    print(call('/api/stock-watch/snapshot', body), f"（{len(body['codes'])} 個代碼，查無 {len(body['missing'])} 個）")


if __name__ == '__main__':
    main()
