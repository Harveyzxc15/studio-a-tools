#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主機 3PP 搭售率趨勢 — 從 DSS「3PP搭售率報表（店）」撈每週、每月，推到 sa-north-hub（/attach-3pp）
  來源＝DSS JSON API（dss_api.fetch '3PP_aggregate_rptD'），北一 dept 01、北二 dept 02
  期間＝週（週日～週六，2026 年起）、月（2025-01 起，去年同期用）、本月累計到昨天＋去年同期同天數
  每店每品類存 搭售量 m／配件 s／零搭售量 ms，比率在頁面算（合計＝各品類加總再除）
  人員別＝「3PP搭售率報表(人)」3PP_aggregate_E，同樣每期每店每人每品類存 m／s／ms
  週報用：另存最近完整週（週日～週六）的 DSS 原始匯出檔（加密，build_deck.py 會自動抓）
  已結束超過 14 天的期間存在快取不再重撈（DSS 銷退會回頭改數字，近兩週每次重撈）

用法：python3 3PP搭售率_推雲端.py            # 撈缺的期間＋推
      python3 3PP搭售率_推雲端.py --weekly   # 排程用（每週一 13:00）：session 過期就跳通知＋開 DSS 登入視窗，登入後接著推；其他錯誤也跳通知
      python3 3PP搭售率_推雲端.py --dry-run  # 只撈不推
      HUB_URL=http://127.0.0.1:5050 python3 3PP搭售率_推雲端.py
⚠️ DSS 登入要圖形驗證碼＋email 驗證碼，session 過期先跑：python3 ../dss_3pp_export.py --login
Token：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token（與 EPB 實績推送共用）
"""
import json, os, ssl, subprocess, sys, urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import dss_api  # noqa: E402
from dss_api import fetch  # noqa: E402

HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
CACHE = Path.home() / '工具中心' / '輸出' / 'DSS3PP' / '趨勢快取.json'
DEPTS = ('01', '02')
CATS = ('cpu', 'iphone', 'ipad', 'watch', 'airpods')
WEEK_FROM = date(2025, 12, 28)     # 含 2026-01-01 的那一週（週日起）
MONTH_FROM = date(2025, 1, 1)      # 去年同期要用到 2025
STABLE_DAYS = 14
DECK_DIR = Path.home() / '工具中心' / '輸出' / 'DSS3PP' / '週報用'
DEPT_NAME = {'01': '北一區', '02': '北二區'}


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


def month_end(d):
    return (d.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)


def periods(today):
    """[(id, kind, start, end)]；資料只到昨天。"""
    last = today - timedelta(days=1)
    out = []
    d = WEEK_FROM
    while d <= last:
        out.append(('W' + d.isoformat(), 'week', d, min(d + timedelta(days=6), last)))
        d += timedelta(days=7)
    d = MONTH_FROM
    while d <= last:
        out.append(('M' + d.strftime('%Y-%m'), 'month', d, min(month_end(d), last)))
        d = month_end(d) + timedelta(days=1)
    # 本月還沒結束：去年同月只取同樣天數，才是同期比較
    if last < month_end(last):
        ly = last.replace(year=last.year - 1)
        out.append(('L' + ly.strftime('%Y-%m'), 'month_ly', ly.replace(day=1), ly))
    return out


def counts(r):
    for c in CATS:
        m, s, ms = (int(float(r.get(f'{c}_{k}') or 0)) for k in ('m', 's', 'ms'))
        if m or s or ms:
            yield c, m, s, ms


def pull(start, end):
    """{'rows': [[店, 品類, m, s, ms]], 'emp': [[店, 員編, 品類, m, s, ms]], 'names': {員編: 名字}}"""
    rows, emp, names = [], [], {}
    for dept in DEPTS:
        for r in fetch('3PP_aggregate_rptD', start.isoformat(), end.isoformat(), dept=dept):
            rows += [[r['SHOP_ID'], *x] for x in counts(r)]
        for r in fetch('3PP_aggregate_E', start.isoformat(), end.isoformat(), dept=dept):
            names[r['emp_id1']] = r.get('emp_name') or r['emp_id1']
            emp += [[r['SHOP_ID'], r['emp_id1'], *x] for x in counts(r)]
    return {'rows': rows, 'emp': emp, 'names': names}


def save_deck_file(today):
    """週報用：最近完整週的 DSS 店別報表原始匯出檔（加密，密碼＝DSS 帳號），兩區各一份。"""
    from dss_3pp_export import periods as last_week
    ws, we, _, _ = last_week(today)
    s = dss_api._session()
    DECK_DIR.mkdir(parents=True, exist_ok=True)
    for dept, name in DEPT_NAME.items():
        where = [{'code': 'start_doc_date', 'value': ws.isoformat()}, {'code': 'end_doc_date', 'value': we.isoformat()},
                 {'code': 'org_id', 'value': '01'}, {'code': 'dept_id', 'value': dept}]
        r = s.get(f'{dss_api.BASE}/rpt/export/3PP_aggregate_rptD', params={'whereColumn': json.dumps(where)}, timeout=120)
        r.raise_for_status()
        if not r.content.startswith(b'\xd0\xcf\x11\xe0'):
            raise SystemExit('❌ session 已失效（匯出週報檔時），請重新執行：dss_3pp_export.py --login')
        path = DECK_DIR / f'3PP_aggregate_rptD_{name}_{ws:%Y%m%d}-{we:%Y%m%d}.xlsx'
        path.write_bytes(r.content)
        print(f'週報用 ✅ {path.name}')


def notify(msg):
    subprocess.run(['osascript', '-e', f'display notification "{msg}" with title "3PP 搭售率" sound name "Glass"'])


def main():
    if '--weekly' not in sys.argv:
        return run()
    try:
        run()
    except Exception:
        # 網路、hub、DSS 錯誤：不跳通知的話使用者不會知道這週沒更新
        notify('更新失敗，看 ~/Library/Logs/attach3pp.log')
        raise
    except SystemExit as e:
        if 'session' not in str(e):
            notify('更新失敗，看 ~/Library/Logs/attach3pp.log')
            raise
        # DSS 登入要圖形驗證碼＋email 驗證碼，只能請使用者自己登入（10 分鐘內）
        notify('DSS 登入過期了，請在跳出的視窗登入，登入後會自動更新')
        from dss_3pp_export import login
        try:
            login()
        except Exception:
            notify('10 分鐘內沒有登入，這週沒更新；之後手動跑 3PP搭售率_推雲端.py')
            raise
        try:
            run()
        except BaseException:
            notify('登入後更新失敗，看 ~/Library/Logs/attach3pp.log')
            raise
    notify('已更新 sa-north-hub 的 3PP 搭售率')


def run():
    today = date.today()
    cache = json.loads(CACHE.read_text()) if CACHE.exists() else {}
    plist = periods(today)
    # 'emp' 是 11/4 加的人員別：舊快取沒有就重撈
    todo = [p for p in plist
            if not (p[0] in cache and cache[p[0]]['end'] == p[3].isoformat() and 'emp' in cache[p[0]]
                    and p[3] <= today - timedelta(days=STABLE_DAYS))]
    print(f'期間 {len(plist)} 個，要撈 {len(todo)} 個')
    # 先單撈一次：session 失效時 dss_api 會直接結束，不會開一堆執行緒一起報錯
    if todo:
        first = todo[0]
        cache[first[0]] = {'end': first[3].isoformat(), **pull(first[2], first[3])}
    with ThreadPoolExecutor(4) as ex:
        for p, got in zip(todo[1:], ex.map(lambda p: pull(p[2], p[3]), todo[1:])):
            cache[p[0]] = {'end': p[3].isoformat(), **got}
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(cache, ensure_ascii=False))
    save_deck_file(today)

    body = {'through': (today - timedelta(days=1)).isoformat(),
            'periods': [{'id': i, 'kind': k, 'start': s.isoformat(), 'end': e.isoformat()}
                        for i, k, s, e in plist],
            'rows': [[i] + r for i, *_ in plist for r in cache[i]['rows']],
            'emp_rows': [[i] + r for i, *_ in plist for r in cache[i]['emp']],
            'emp_names': {k: v for i, *_ in plist for k, v in cache[i]['names'].items()}}
    if '--dry-run' in sys.argv:
        print(json.dumps({k: len(v) if isinstance(v, (list, dict)) else v for k, v in body.items()}))
        return
    print(call('/api/attach3pp/snapshot', body))


if __name__ == '__main__':
    main()
