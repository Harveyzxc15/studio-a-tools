#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主機 3PP 搭售率趨勢 — 從 DSS「3PP搭售率報表（店）」撈每週、每月，推到 sa-north-hub（/attach-3pp）
  來源＝DSS JSON API（dss_api.fetch '3PP_aggregate_rptD'），北一 dept 01、北二 dept 02
  期間＝週（週日～週六，2026 年起）、月（2025-01 起，去年同期用）、本月累計到昨天＋去年同期同天數
  每店每品類存 搭售量 m／配件 s／零搭售量 ms，比率在頁面算（合計＝各品類加總再除）
  已結束超過 14 天的期間存在快取不再重撈（DSS 銷退會回頭改數字，近兩週每次重撈）

用法：python3 3PP搭售率_推雲端.py            # 撈缺的期間＋推
      python3 3PP搭售率_推雲端.py --dry-run  # 只撈不推
      HUB_URL=http://127.0.0.1:5050 python3 3PP搭售率_推雲端.py
⚠️ DSS 登入要圖形驗證碼＋email 驗證碼，session 過期先跑：python3 ../dss_3pp_export.py --login
Token：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token（與 EPB 實績推送共用）
"""
import json, os, ssl, sys, urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dss_api import fetch  # noqa: E402

HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
CACHE = Path.home() / '工具中心' / '輸出' / 'DSS3PP' / '趨勢快取.json'
DEPTS = ('01', '02')
CATS = ('cpu', 'iphone', 'ipad', 'watch', 'airpods')
WEEK_FROM = date(2025, 12, 28)     # 含 2026-01-01 的那一週（週日起）
MONTH_FROM = date(2025, 1, 1)      # 去年同期要用到 2025
STABLE_DAYS = 14


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


def pull(start, end):
    rows = []
    for dept in DEPTS:
        for r in fetch('3PP_aggregate_rptD', start.isoformat(), end.isoformat(), dept=dept):
            for c in CATS:
                m, s, ms = (int(float(r.get(f'{c}_{k}') or 0)) for k in ('m', 's', 'ms'))
                if m or s or ms:
                    rows.append([r['SHOP_ID'], c, m, s, ms])
    return rows


def main():
    today = date.today()
    cache = json.loads(CACHE.read_text()) if CACHE.exists() else {}
    plist = periods(today)
    todo = [p for p in plist
            if not (p[0] in cache and cache[p[0]]['end'] == p[3].isoformat()
                    and p[3] <= today - timedelta(days=STABLE_DAYS))]
    print(f'期間 {len(plist)} 個，要撈 {len(todo)} 個')
    # 先單撈一次：session 失效時 dss_api 會直接結束，不會開一堆執行緒一起報錯
    if todo:
        first = todo[0]
        cache[first[0]] = {'end': first[3].isoformat(), 'rows': pull(first[2], first[3])}
    with ThreadPoolExecutor(4) as ex:
        for p, rows in zip(todo[1:], ex.map(lambda p: pull(p[2], p[3]), todo[1:])):
            cache[p[0]] = {'end': p[3].isoformat(), 'rows': rows}
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(cache, ensure_ascii=False))

    body = {'through': (today - timedelta(days=1)).isoformat(),
            'periods': [{'id': i, 'kind': k, 'start': s.isoformat(), 'end': e.isoformat()}
                        for i, k, s, e in plist],
            'rows': [[i] + r for i, *_ in plist for r in cache[i]['rows']]}
    if '--dry-run' in sys.argv:
        print(json.dumps({'through': body['through'], 'periods': len(body['periods']), 'rows': len(body['rows'])}))
        return
    print(call('/api/attach3pp/snapshot', body))


if __name__ == '__main__':
    main()
