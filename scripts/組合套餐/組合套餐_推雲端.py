#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
iPhone 配件組合 — 本機查 EPB，每日各店組數推到 sa-north-hub（/combo）
  組數＝poslinev_bi 北區 org 01、stk_id 99200201「促銷組合」列的 stk_qty，依 PBCAM_DOC_ID 分組合
        trans_type A/E/G/J（訂金單也有組合列；不含 H 免得訂金＋尾款重複算）
  iPhone＝cat4 4004、trans_type A/E/H 的 stk_qty
  範圍＝檔期起（START）～今天，依日×店存，期間加總在頁面做

用法：python3 組合套餐_推雲端.py            # 直接推
      python3 組合套餐_推雲端.py --poll     # 有人按「刷新」，或今天 08:30 後還沒推過才推（launchd 每 2 分鐘跑）
      python3 組合套餐_推雲端.py --dry-run  # 只查不推
      HUB_URL=http://127.0.0.1:5050 python3 組合套餐_推雲端.py
備援機（筆電 launchd 設 SYNC_STANDBY=1）：請求掛滿 3 分鐘、mini 都沒推才接手
Token：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token（與 EPB 實績推送共用）
換檔期：改 START 與 COMBOS（活動代碼查 pos_mc_mas app_code='POSCAM2N'）
"""
import csv, io, json, os, ssl, subprocess, sys, urllib.request
from datetime import date, datetime

HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
JAVA = '/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java'
LIB = os.path.expanduser('~/工具中心/lib')          # EPBReportQuery.class（不放 Desktop：launchd 讀不到）
EPB_CP = f'{LIB}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*'
# 13 店（與 sa_north_hub/stores.py 同一份）；mini 上沒有 git，自含一份免得 import 不到
STORES = ['004', '005', '046', '024', '054', '057', '068', '063', '025', '050', '009', '055', '064']
START = '2026-09-15'                               # 檔期 T0926_S
COMBOS = [  # POSCAM2N 活動代碼、頁面欄名、加購價
    {'id': '099260914001', 'label': 'A 殼+貼+拭鏡紙', 'price': 2280},
    {'id': '099260914005', 'label': 'A 升級 Simmpo', 'price': 2580},
    {'id': '099260914003', 'label': 'B 貼+鏡頭貼+拭鏡紙', 'price': 1680},
    {'id': '099260914006', 'label': 'B 升級 Simmpo', 'price': 1980},
    {'id': '099260914002', 'label': 'C 殼+鏡頭貼+拭鏡紙', 'price': 1580},
    {'id': '099260914004', 'label': 'D 45W+傳輸線', 'price': 1280},
]
DAILY_AFTER = '08:30'                              # 每天這個時間後第一次輪詢自動推


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


STANDBY_STATE = os.path.expanduser('~/.config/studioa/combo_standby_seen.json')


def standby_ready(due):
    """同 庫存追蹤_推雲端.py：備援機等該推的狀態掛滿 3 分鐘才做；主機（沒設 SYNC_STANDBY）一律直接做。"""
    if os.environ.get('SYNC_STANDBY') != '1':
        return due
    if not due:
        if os.path.exists(STANDBY_STATE):
            os.remove(STANDBY_STATE)
        return due
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


def fetch(today):
    base = (f"FROM poslinev_bi l WHERE l.org_id='01' AND l.shop_id IN {in_list(STORES)}"
            f" AND l.doc_date>=TO_DATE('{START}','YYYY-MM-DD')")
    sets = [[r['D'], r['SHOP_ID'].zfill(3), r['PBCAM_DOC_ID'], int(float(r['Q']))]
            for r in run_sql(f"SELECT TO_CHAR(l.doc_date,'YYYY-MM-DD') d, l.shop_id, l.pbcam_doc_id, SUM(l.stk_qty) q {base}"
                             f" AND l.stk_id='99200201' AND l.pbcam_code='POSCAM2N'"
                             f" AND l.pbcam_doc_id IN {in_list(c['id'] for c in COMBOS)}"
                             f" AND l.trans_type IN ('A','E','G','J')"
                             f" GROUP BY TO_CHAR(l.doc_date,'YYYY-MM-DD'), l.shop_id, l.pbcam_doc_id")]
    iphone = [[r['D'], r['SHOP_ID'].zfill(3), int(float(r['Q']))]
              for r in run_sql(f"SELECT TO_CHAR(l.doc_date,'YYYY-MM-DD') d, l.shop_id, SUM(l.stk_qty) q {base}"
                               f" AND l.cat4_id='4004' AND l.trans_type IN ('A','E','H')"
                               f" GROUP BY TO_CHAR(l.doc_date,'YYYY-MM-DD'), l.shop_id")]
    return {'start': START, 'end': str(today), 'combos': COMBOS,
            'sets': [s for s in sets if s[3]], 'iphone': [i for i in iphone if i[2]]}


def main():
    pending = call('/api/combo/pending')
    if '--poll' in sys.argv:
        now = datetime.now()
        daily_due = pending.get('pushed_on') != str(now.date()) and now.strftime('%H:%M') >= DAILY_AFTER
        if not standby_ready(pending.get('pending') or daily_due):
            return
    body = fetch(date.today())
    if '--dry-run' in sys.argv:
        print(json.dumps({k: v if k in ('start', 'end') else len(v) for k, v in body.items()}, ensure_ascii=False))
        return
    print(call('/api/combo/snapshot', body), f"（{sum(s[3] for s in body['sets'])} 組）")


if __name__ == '__main__':
    main()
