#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
新品追蹤 — 本機查 EPB，各追蹤機型的「日×店」台數、AC+、SA Care 推到 sa-north-hub（/new-products）
  追蹤哪些機型＝網站上管理員設定（名稱、cat6、首賣日），這支從 /api/new-products/pending 取清單
  台數＝poslinev_bi 北區 org 01、該機型 cat6、trans_type A/E/H 的 stk_qty（銷售＋尾款－銷退，不含訂金 G）
  AC+＝「含該機型的單」裡 cat3 3032、品名符合該產品線的列（同記憶 feedback_warranty_attachment_rule）
        （不用品名關鍵字；Duo、iPad mini A20、MBP 14 M6 到 11/10 逐店對過之前交付的數字一致）
  SA Care＝SACare cat6（6533～6537）且品名含該機型設定的關鍵字（例：iPhone Duo），不限同單
        （現在幾乎都是「檢測新機」，跟主機不同單，所以用品名認；沒設關鍵字就不算）
  ⚠️ 同一張單有兩台同產品線不同機型時，AC+ 會算到被追蹤的那台；保固單獨退（退貨單沒有主機）不會扣
  維修中心 110→士林 004、113→花蓮 055
  另推「近 90 天首次開賣的主機 cat6」當候選清單，網站上新增機型時直接挑

用法：python3 新品追蹤_推雲端.py            # 直接推
      python3 新品追蹤_推雲端.py --poll     # 有人按「刷新」，或今天 08:30 後還沒推過才推（launchd 每 2 分鐘跑）
      python3 新品追蹤_推雲端.py --dry-run  # 只查不推
      HUB_URL=http://127.0.0.1:5050 python3 新品追蹤_推雲端.py
Token：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token（與其他推送共用）
"""
import csv, io, json, os, ssl, subprocess, sys, urllib.request
from datetime import date, datetime, timedelta

HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
JAVA = '/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java'
LIB = os.path.expanduser('~/工具中心/lib')
EPB_CP = f'{LIB}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*'
STORES = ['004', '005', '046', '024', '054', '057', '068', '063', '025', '050', '009', '055', '064']
REPAIR = {'110': '004', '113': '055'}              # 維修中心併回門市
DAILY_AFTER = '08:30'
# 產品線：主機 cat4、AC+ 品名關鍵字
LINES = {
    'Mac': ({'4001', '4002'}, ('macbook', 'mac mini', 'imac', 'mac studio')),
    'iPad': ({'4005', '4006', '4041'}, ('ipad',)),
    'iPhone': ({'4004'}, ('iphone',)),
    'Watch': ({'4038'}, ('watch',)),
    'AirPods': ({'4014'}, ('airpods',)),
}
HOST_CAT4 = sorted(c for cat4, _ in LINES.values() for c in cat4)
SACARE_CAT6 = ['6533', '6534', '6535', '6536', '6537']


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


def run_sql(sql, max_rows=100000):
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


def aggregate(lines, cat6, sacare=()):
    """含該機型的單的明細列＋SA Care 列 [(日, 店, 數量)] → (產品線, [[日, 店, 台數, AC+, SACare], ...])，全 0 不列。"""
    cat6 = set(cat6)
    line = next((name for l in lines if l['CAT6_ID'] in cat6
                 for name, (cat4, _) in LINES.items() if l['CAT4_ID'] in cat4), None)
    words = LINES[line][1] if line else ()
    out = {}

    def cell(d, shop):
        shop = shop.zfill(3)
        return out.setdefault((d, REPAIR.get(shop, shop)), [0, 0, 0])

    for l in lines:
        q = int(float(l['Q']))
        if l['CAT6_ID'] in cat6:
            cell(l['D'], l['SHOP_ID'])[0] += q
        elif l['CAT3_ID'] == '3032' and any(w in l['NAME'].lower() for w in words):
            cell(l['D'], l['SHOP_ID'])[1] += q
    for d, shop, q in sacare:
        cell(d, shop)[2] += q
    return line, [[d, s, *v] for (d, s), v in sorted(out.items()) if any(v)]


def fetch_model(model):
    shops = in_list(STORES + list(REPAIR))
    scope = (f"org_id='01' AND shop_id IN {shops} AND doc_date>=TO_DATE('{model['start']}','YYYY-MM-DD')"
             f" AND trans_type IN ('A','E','H')")
    lines = run_sql(
        f"SELECT TO_CHAR(l.doc_date,'YYYY-MM-DD') d, l.shop_id, l.cat3_id, l.cat4_id, l.cat6_id, l.stk_qty q, l.name"
        f" FROM poslinev_bi l WHERE l.{scope.replace(' AND ', ' AND l.')}"
        f" AND (l.cat6_id IN {in_list(model['cat6'])} OR l.cat3_id='3032')"
        f" AND l.doc_id IN (SELECT x.doc_id FROM poslinev_bi x WHERE x.{scope.replace(' AND ', ' AND x.')}"
        f" AND x.cat6_id IN {in_list(model['cat6'])})")
    sacare = []
    kw = (model.get('sacare_kw') or '').strip().replace("'", "''")
    if kw:
        sacare = [(r['D'], r['SHOP_ID'], int(float(r['Q']))) for r in run_sql(
            f"SELECT TO_CHAR(l.doc_date,'YYYY-MM-DD') d, l.shop_id, SUM(l.stk_qty) q FROM poslinev_bi l"
            f" WHERE l.{scope.replace(' AND ', ' AND l.')} AND l.cat6_id IN {in_list(SACARE_CAT6)}"
            f" AND l.name LIKE '%{kw}%' GROUP BY TO_CHAR(l.doc_date,'YYYY-MM-DD'), l.shop_id")]
    line, rows = aggregate(lines, model['cat6'], sacare)
    return {'id': model['id'], 'line': line, 'rows': rows}


def fetch_candidates(today):
    """近 90 天才第一次賣的主機 cat6（往前看一年確認之前沒賣過），排除 SACare 類別 3012。"""
    rows = run_sql(
        f"SELECT l.cat4_id, l.cat6_id, MAX(l.name) nm, TO_CHAR(MIN(l.doc_date),'YYYY-MM-DD') d0, SUM(l.stk_qty) q"
        f" FROM poslinev_bi l WHERE l.org_id='01' AND l.doc_date>=TO_DATE('{today - timedelta(days=365)}','YYYY-MM-DD')"
        f" AND l.trans_type IN ('A','E','H') AND l.cat4_id IN {in_list(HOST_CAT4)} AND l.cat3_id<>'3012'"
        f" GROUP BY l.cat4_id, l.cat6_id"
        f" HAVING MIN(l.doc_date)>=TO_DATE('{today - timedelta(days=90)}','YYYY-MM-DD') ORDER BY 4 DESC", 2000)
    return [{'cat6': r['CAT6_ID'], 'name': r['NM'][:60], 'start': r['D0'], 'qty': int(float(r['Q']))} for r in rows]


def epb_up():
    import socket
    try:
        with socket.create_connection(('192.168.1.177', 8080), timeout=1.5):
            return True
    except OSError:
        return False


def main():
    pending = call('/api/new-products/pending')
    if '--poll' in sys.argv:
        now = datetime.now()
        daily_due = pending.get('pushed_on') != str(now.date()) and now.strftime('%H:%M') >= DAILY_AFTER
        if not (pending.get('pending') or daily_due) or not epb_up():
            return
    today = date.today()
    body = {'end': str(today), 'models': [fetch_model(m) for m in pending.get('models', [])],
            'candidates': fetch_candidates(today)}
    if '--dry-run' in sys.argv:
        for m in body['models']:
            print(m['id'], m['line'], '台數', sum(r[2] for r in m['rows']), 'AC+', sum(r[3] for r in m['rows']),
                  'SACare', sum(r[4] for r in m['rows']))
        print('候選', len(body['candidates']))
        return
    print(call('/api/new-products/snapshot', body), f"（{len(body['models'])} 個機型）")


if __name__ == '__main__':
    main()
