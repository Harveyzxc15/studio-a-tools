#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主機 3PP 零搭教練 — Mac mini 每天 08:00 用 EPB 重建 DSS 搭售貼標、拆零搭原因，推 sa-north-hub
（/attach-3pp「零搭教練」分頁）＋ Discord（預警每天、行動清單週一）
  期間＝最近 12 個完整週（週日～週六）＋本週到昨天，北一＋北二 13 店
  零搭原因（每台零搭主機一個）：
    計數規則造成，不是銷售問題：同單多台（同單同品類已有一台算搭售）／配件歸給較貴主機（同單有配件，
      但歸到更貴的別品類主機）／同客戶他單已搭售（同客戶同日同店另一張單已算這品類搭售）
    大量單：同單同品類 3 台以上都沒配件（整批採購／代購），配件本來就難賣，也不列入檢討
    真零搭依客人展現的消費訊號分級：A 已買保固（SA Care 3012、AppleCare／AC+ 3032）／B 買了原廠配件
      （3002 非 AirPods）／C 認證機或福利機（主機 3033／3035）／D 有其他訊號（教育價 3006、會員點數 3027）／E 完全沒有訊號
  預警、行動清單、追蹤成效都在 hub 算（attach3pp.py），這裡只負責 EPB 取數與推送
  ⚠️ EPB 重建是從 DSS 反推的規則：北一與 DSS 100% 一致，北二每週約 1–3 台不同（記憶 dss-3pp-epb-rebuild）

用法：python3 3PP預警_推雲端.py              # 算＋推 hub＋Discord（週一另推行動清單）
      python3 3PP預警_推雲端.py --dry-run    # 只算不推，印摘要
      python3 3PP預警_推雲端.py --no-discord # 推 hub 但不推 Discord
      HUB_URL=http://127.0.0.1:5050 python3 3PP預警_推雲端.py
Token：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token；Discord：~/.config/studioa/attach3pp_discord_webhook.txt
"""
import json, os, ssl, sys, urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import dss_rebuild_from_epb as rb  # noqa: E402

HUB = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev').rstrip('/')
WEBHOOK = Path.home() / '.config' / 'studioa' / 'attach3pp_discord_webhook.txt'
WEEKS = 12                                     # 管制圖基準＝前 12 個完整週
DETAIL_WEEKS = 2                               # 零搭明細只帶最近兩週（上週＋本週），給店長跟夥伴回顧
CATS = r'^(CPU|iPhone|iPad|Watch|Airpods)'
BULK = 3                                       # 同單同品類真零搭幾台以上算大量單


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
    with urllib.request.urlopen(req, timeout=120, context=context) as r:
        return json.loads(r.read())


def discord(text):
    req = urllib.request.Request(WEBHOOK.read_text().strip(), data=json.dumps({'content': text}).encode(),
                                 headers={'Content-Type': 'application/json', 'User-Agent': 'sa-north-hub'})
    try:
        import certifi
        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        context = None
    urllib.request.urlopen(req, timeout=60, context=context).read()


def sunday(d):
    return d - timedelta(days=(d.weekday() + 1) % 7)


def marks(start, end, shops):
    """單據上的消費訊號（只撈需要的類別，全品類 13 週會超過筆數上限）。"""
    shop_in = ','.join(f"'{s}'" for s in shops)
    m = rb.epb(
        "SELECT DISTINCT doc_id, cat3_id FROM poslinev_bi WHERE org_id='01'"
        f" AND shop_id IN ({shop_in}) AND doc_date BETWEEN to_date('{start}','YYYY-MM-DD')"
        f" AND to_date('{end}','YYYY-MM-DD')"
        " AND (cat3_id IN ('3012','3032','3006','3027') OR (cat3_id='3002' AND cat4_id<>'4014'))")
    return {'warranty': set(m[m.CAT3_ID.isin(['3012', '3032'])].DOC_ID),
            'orig': set(m[m.CAT3_ID == '3002'].DOC_ID),
            'other': set(m[m.CAT3_ID.isin(['3006', '3027'])].DOC_ID)}


def tag(d, mk):
    """貼標後的明細 → 加品類、種類（m 搭售／s 配件／ms 零搭）、零搭原因、數量。"""
    d = d.copy()
    d['qty'] = pd.to_numeric(d.STK_QTY, errors='coerce').fillna(0)
    d['amt'] = pd.to_numeric(d.LINE_TOTAL, errors='coerce').fillna(0)
    d['cat'] = d.type_name.str.extract(CATS, expand=False)
    d['kind'] = None
    d.loc[d.type_name.str.endswith('配件'), 'kind'] = 's'
    d.loc[d.type_name.str.endswith('搭售'), 'kind'] = 'm'
    d.loc[d.type_name.str.contains('零搭售'), 'kind'] = 'ms'
    att = d[d.kind == 'm']
    by_doc = att.groupby('DOC_ID').cat.agg(set).to_dict()
    # 跨單限同一家店（同 dss_rebuild_from_epb 的規則）
    by_vip = att[att.VIP_ID.notna()].groupby(['VIP_ID', 'DOC_DATE', 'SHOP_ID']).cat.agg(set).to_dict()

    def reason(r):
        same = by_doc.get(r.DOC_ID, set())
        if r.cat in same:
            return '同單多台'
        if same:
            return '配件歸給較貴主機'
        if pd.notna(r.VIP_ID) and r.cat in by_vip.get((r.VIP_ID, r.DOC_DATE, r.SHOP_ID), set()):
            return '同客戶他單已搭售'
        if r.DOC_ID in mk['warranty']:
            return 'A'
        if r.DOC_ID in mk['orig']:
            return 'B'
        if r.CAT3_ID in ('3033', '3035'):
            return 'C'
        if r.DOC_ID in mk['other']:
            return 'D'
        return 'E'

    zero = d.kind == 'ms'
    d['reason'] = ''
    d.loc[zero, 'reason'] = [reason(r) for r in d[zero].itertuples()]
    # 大量單：同一張單同品類 3 台以上全是真零搭（11/3 新西門一張單 6 台 iPhone 18 Pro Max，0 配件）
    real = d.reason.isin(list('ABCDE'))
    big = d[real].groupby(['DOC_ID', 'cat']).qty.sum()
    big = set(big[big >= BULK].index)
    d.loc[real & pd.Series([(a, b) in big for a, b in zip(d.DOC_ID, d.cat)], index=d.index), 'reason'] = '大量單'
    return d


def build(today):
    last = today - timedelta(days=1)
    cur = sunday(last)
    start = cur - timedelta(days=7 * WEEKS)
    parts = []
    for dept, shops in rb.REGIONS.items():
        lines, occ = rb.pull(start.isoformat(), last.isoformat(), shops)
        parts.append(tag(rb.label(lines, occ), marks(start.isoformat(), last.isoformat(), shops)))
    d = pd.concat(parts, ignore_index=True)
    d['week'] = [sunday(date.fromisoformat(x)).isoformat() for x in d.DOC_DATE]

    g = (d.groupby(['week', 'SHOP_ID', 'EMP_ID1', 'cat', 'kind', 'reason'], dropna=False).qty.sum()
         .reset_index())
    g = g[g.qty != 0]
    rows = [[r.week, r.SHOP_ID, r.EMP_ID1 if pd.notna(r.EMP_ID1) else '', r.cat, r.kind, r.reason, int(round(r.qty))]
            for r in g.itertuples()]

    recent = (cur - timedelta(days=7 * (DETAIL_WEEKS - 1))).isoformat()
    z = d[(d.week >= recent) & d.reason.isin(list('ABCDE'))].sort_values(['DOC_DATE', 'DOC_ID'])
    detail = [[r.DOC_DATE, f'{r.DOC_ID.split("-")[2][8:10]}:{r.DOC_ID.split("-")[2][10:12]}' if r.DOC_ID.count('-') >= 2 else '',
               r.SHOP_ID, r.EMP_ID1 if pd.notna(r.EMP_ID1) else '', r.cat, r.NAME, int(round(r.amt)),
               int(round(r.qty)), r.reason, r.DOC_ID] for r in z.itertuples()]

    weeks = [{'id': (start + timedelta(days=7 * i)).isoformat(), 'start': (start + timedelta(days=7 * i)).isoformat(),
              'end': min(start + timedelta(days=7 * i + 6), last).isoformat()} for i in range(WEEKS + 1)]
    return {'today': today.isoformat(), 'through': last.isoformat(), 'weeks': weeks, 'rows': rows, 'detail': detail}


def main():
    today = date.today()
    body = build(today)
    if '--dry-run' in sys.argv:
        ms = sum(r[6] for r in body['rows'] if r[4] == 'ms')
        rule = sum(r[6] for r in body['rows'] if r[4] == 'ms' and r[5] not in 'ABCDE')
        print(json.dumps({'through': body['through'], 'weeks': len(body['weeks']), 'rows': len(body['rows']),
                          'detail': len(body['detail']), '零搭': ms, '計數規則造成': rule}, ensure_ascii=False))
        return
    res = call('/api/attach3pp/epb', body)
    print(datetime.now().strftime('%m-%d %H:%M'), {k: v for k, v in res.items() if k != 'discord'})
    if '--no-discord' in sys.argv:
        return
    msgs = []
    if res['discord'].get('alerts'):
        msgs.append(res['discord']['alerts'])
    if today.weekday() == 0 or '--actions' in sys.argv:      # 週一推行動清單
        msgs.append(res['discord']['actions'])
    for text in msgs:
        discord(text)
    print(f'Discord 推 {len(msgs)} 則')


if __name__ == '__main__':
    main()
