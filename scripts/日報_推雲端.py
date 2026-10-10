#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
門市日報 — 用 EPB 重算 13 店日報數字，推到 sa-north-hub
======================================
雲端打不到 EPB（只走 VPN），所以本機每天查一次、POST 上去，
sa-north-hub 的 /daily 頁照日報格式把 13 店並排。

口徑照公司日報（北一區_北一區日報_YYYYMMDD.xlsx 的 sell／ttl／sac 分頁），
2026-10-03 用 13 店 10/1、10/2 日報逐格對過：
  零售營業額(a)＝A＋E 金額　　訂單出貨(b)＝H　　零售訂金(c)＝G＋J 實收實退的訂金（deposit 欄）
  Apple 主機／Apple 配件／3PP／AC+ 金額＝cat3 3001／3002／3003、cat4 4969 的 A+E+G+J（J 用實退訂金）
  成交筆數＝銷售單＋訂單出貨－退貨單，退貨單只扣「原單也在同一期間」的
  台數、保固套數＝A+E+H 數量（AC+ cat6 6524-6529、SA Care cat6 6533-6537）
  SA Care 金額＝BIPOS_VIEW_BI cat3 3012／cat2 2002 的 A/E/H 數量 × 價目表（EPB 金額是 0）
  毛利＝BI「06／07／08-日報 毛利額 含稅（排除銀聯）」，Apple 另外扣掉 AC+（日報 Apple 毛利不含 AC+）
  AC+ 毛利＝營業額 − 標準成本×1.05（成本隔天才補上，所以當天推的會偏高，月累積會自己修正）

【用法】
    python3 日報_推雲端.py                 # 推昨天（launchd 每天早上跑）
    python3 日報_推雲端.py 2026-10-02      # 推指定日期
    python3 日報_推雲端.py 2026-10-02 --dry  # 只印不推
    HUB_URL=http://127.0.0.1:5000 python3 日報_推雲端.py   # 測本機

【Token】同 目標分配_推雲端.py：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token
======================================
"""
import json
import os
import sys
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import 目標分配_填業績 as F
import 目標分配_推雲端 as P

LOCS = F.NORTH1 + F.NORTH2
LOC_IN = ','.join(f"'{loc}'" for loc in LOCS)

# 「排除銀聯」＝排除這些 VIP（日報 BI 的 VIP_ID 篩選；enqsales_bi 這邊叫 cust_id，清單在 F.CUST）
VIP_EXCL = ("'SA03139209','ST00237177','SA03189128','SA03147688','SA03125724','SA03132099',"
            "'SA03145281','SA03129217','SA03130573','SA03175171','SA03147048','SA03184320'")
# 毛利 BI 06／07 的 STK_ID 排除清單（08 3PP 用後面那段較短的）
GM_STK_EXCL = ("'07307154','07309136','07309137','07310037','07310042','07310053','07310093',"
               "'07311242','88600895','88601027','90501795','90501799','99200168','99500006',"
               "'99900946','99900947','99900948','99900949','99900950','99901684','99901685',"
               "'99902607','99902608','99902609','99902610','99903343','99903339'")
TPP_STK_EXCL = ("'99200168','99500006','99900946','99900947','99900948','99900949','99900950',"
                "'99901684','99901685','99902607','99902608','99902609','99902610'")
GM_COMMON = ("(cat6_id not in ('6888','6889') or cat6_id is null) "
             "and (cat1_id not in ('1002','1004','1008') or cat1_id is null) "
             "and brand_id not in ('297') and (class_id not in ('05') or class_id is null) "
             "and (trans_type not in ('G','J') or trans_type is null) "
             f"and (vip_id not in ({VIP_EXCL}) or vip_id is null)")

# 保固套數：AC+ 與 SA Care 的 cat6，依品類
WARRANTY = {'cpu': ('6524', '6533'), 'ipad': ('6526', '6534'), 'iphone': ('6525', '6535'),
            'watch': ('6528', '6536'), 'air': ('6529', '6537')}


def _date_range(d1, d2):
    return (f"doc_date>=to_date('{d1.isoformat()}','yyyy-mm-dd') "
            f"and doc_date<to_date('{(d2 + timedelta(days=1)).isoformat()}','yyyy-mm-dd')")


def _num(value):
    text = str(value)
    return 0.0 if text in ('nan', 'None', '') else float(text)


def _by_day(df, keys):
    """EPB 回傳（每店每天一列）→ {(loc, 'YYYY-MM-DD'): {key: 值}}，金額先不四捨五入。"""
    out = {}
    if df is None:
        return out
    for _, r in df.iterrows():
        out[(str(r['LOC_ID']).strip(), str(r['D'])[:10])] = {k: _num(r[k.upper()]) for k in keys}
    return out


def query_sell(d1, d2):
    """sell 分頁：金額、台數、保固套數（enqsales_bi，同 目標分配 的單別／客戶排除）。"""
    # 品類金額：退訂(J) 用實退的 deposit（預購單可能只收了部分訂金），其餘用折後金額。
    # 訂金(G) 照日報用全額（13 店 10/1、10/2 對過：微風 AirPods 預購收 2000 訂金，配件仍算 4490）
    paid = "case when trans_type='J' then nvl(deposit,0) else line_total_aftdisc end"
    amt = "trans_type in ('A','E','G','J')"
    def a(cond):
        return f"sum(case when {amt} and {cond} then {paid} else 0 end)"
    def q(cond):
        return f"sum(case when {cond} then stk_qty else 0 end)"
    cols = {
        'retail': "sum(case when trans_type in ('A','E') then line_total_aftdisc else 0 end)",
        'order_ship': "sum(case when trans_type='H' then line_total_aftdisc else 0 end)",
        'deposit': "sum(case when trans_type in ('G','J') then nvl(deposit,0) else 0 end)",
        'apple_host': a("cat3_id='3001'"), 'apple_acc': a("cat3_id='3002'"),
        'tpp': a("cat3_id='3003'"), 'acpp_amt': a("cat4_id='4969'"),
        'cpu': q("cat4_id in ('4001','4002')"),          # 日報 CPU 台數＝CPU＋Mac mini
        'ipad': q("cat4_id in ('4005','4006','4041')"),
        'iphone': q("cat4_id='4004'"), 'watch': q("cat4_id='4038'"),
        'air': q("cat4_id='4014' and cat6_id<>'6070' and cat3_id='3002'"),
    }
    for cat, (acpp, sac) in WARRANTY.items():
        cols[f'{cat}_w'] = q(f"cat6_id in ('{acpp}','{sac}')")
    sql = ("select loc_id, to_char(doc_date,'YYYY-MM-DD') d, "
           + ', '.join(f'{v} {k}' for k, v in cols.items())
           + f" from enqsales_bi where org_id='01' and loc_id in ({LOC_IN}) and {_date_range(d1, d2)}"
           f" and app_code in ('RNCN','DNN','POSN') and (cust_id is null or cust_id not in {F.CUST})"
           " group by loc_id, to_char(doc_date,'YYYY-MM-DD')")
    return _by_day(F.run_epb(sql), cols)


def query_margin(d1, d2):
    """ttl 分頁：總／Apple／3PP 毛利的營收與成本（BIPOS_VIEW_BI，四捨五入留到期間加總後）。"""
    price = "decode(trans_type,'G',0,'I',0,'J',0,'K',0,qty*unit_price)"
    cost = "decode(trans_type,'G',0,'I',0,'J',0,'K',0,qty*unit_cost)*1.05"
    conds = {
        'gm_total': f"(cat3_id not in ('3047') or cat3_id is null) and stk_id not in ({GM_STK_EXCL})",
        'gm_apple': ("(cat3_id not in ('3047','3003','3004','3018','3019','3006') or cat3_id is null) "
                     f"and (cat4_id<>'4969' or cat4_id is null) and stk_id not in ({GM_STK_EXCL})"),
        'gm_tpp': f"cat3_id in ('3003','3004','3018','3019','3006') and stk_id not in ({TPP_STK_EXCL})",
    }
    cols = {}
    for key, cond in conds.items():
        cols[f'{key}_p'] = f"sum(case when {cond} then {price} else 0 end)"
        cols[f'{key}_c'] = f"sum(case when {cond} then {cost} else 0 end)"
    sql = ("select loc_id, to_char(doc_date,'YYYY-MM-DD') d, "
           + ', '.join(f'{v} {k}' for k, v in cols.items())
           + f" from bipos_view_bi where org_id='01' and loc_id in ({LOC_IN}) and {_date_range(d1, d2)}"
           f" and {GM_COMMON} group by loc_id, to_char(doc_date,'YYYY-MM-DD')")
    return _by_day(F.run_epb(sql), cols)


def query_acpp_margin(d1, d2):
    """ttl 分頁 acp：AC+ 營業額 − 標準成本×1.05（std_cost 隔天才補，當天推的會偏高）。"""
    keep = "decode(trans_type,'G',0,'I',0,'J',0,'K',0,"
    sql = ("select loc_id, to_char(doc_date,'YYYY-MM-DD') d, "
           f"sum({keep}line_total_aftdisc)) acp_p, sum({keep}stk_qty*std_cost))*1.05 acp_c "
           f"from enqsales_bi where org_id='01' and loc_id in ({LOC_IN}) and {_date_range(d1, d2)} "
           "and cat4_id='4969' group by loc_id, to_char(doc_date,'YYYY-MM-DD')")
    return _by_day(F.run_epb(sql), ['acp_p', 'acp_c'])


def sacare_prices():
    import pandas as pd
    prices = pd.read_excel(F.SACARE_PRICE_FILE, header=0)
    prices.columns = ['stk_id', 'name', 'price']
    prices['stk_id'] = prices['stk_id'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return prices.dropna(subset=['stk_id']).drop_duplicates('stk_id').set_index('stk_id')['price'].to_dict()


def query_sacare(d1, d2):
    """sac 分頁：SA Care 數量 × 價目表（BI「10-日報製作匯出 SA Care (排除銀聯)」的條件）。"""
    sql = ("select loc_id, to_char(doc_date,'YYYY-MM-DD') d, stk_id, max(name) nm, sum(qty) qty "
           f"from bipos_view_bi where org_id='01' and loc_id in ({LOC_IN}) and {_date_range(d1, d2)} "
           "and cat3_id like '3012' and cat2_id like '2002' and trans_type in ('A','E','H') "
           f"and (vip_id not in ({VIP_EXCL}) or vip_id is null) "
           "group by loc_id, to_char(doc_date,'YYYY-MM-DD'), stk_id")
    df = F.run_epb(sql)
    prices, out, unmatched = sacare_prices(), {}, []
    if df is None:
        return out, unmatched
    for _, r in df.iterrows():
        qty, stk = _num(r['QTY']), str(r['STK_ID']).strip()
        if stk not in prices and qty:
            unmatched.append((str(r['LOC_ID']), stk, r['NM'], qty))
        key = (str(r['LOC_ID']).strip(), str(r['D'])[:10])
        out.setdefault(key, {'sac_amt': 0.0})['sac_amt'] += qty * prices.get(stk, 0)
    return out, unmatched


def query_tradein(d1, d2):
    """tradein 分頁：舊換新台數（BI「日報舊換新檢核」：cat6 6990、排除折抵金額品項、A/E/H 數量）。
    回收品在 EPB 是負數量，日報顯示成正的台數，所以取負號。"""
    sql = ("select loc_id, to_char(doc_date,'YYYY-MM-DD') d, "
           "-sum(decode(trans_type,'A',stk_qty,'E',stk_qty,'H',stk_qty,0)) tradein "
           f"from poslinev_bi where org_id='01' and loc_id in ({LOC_IN}) and {_date_range(d1, d2)} "
           "and cat6_id='6990' and name not like '%折抵金額%' group by loc_id, to_char(doc_date,'YYYY-MM-DD')")
    return _by_day(F.run_epb(sql), ['tradein'])


def query_txn(d1, d2):
    """sellcnt 分頁：成交筆數＝銷售單(A)＋訂單出貨(H)－退貨單(E)，但退貨只扣「退的是同一期間內的單」。

    13 店 10/1、10/2 對過：退前幾天賣的單，今日不扣；月累積則只要原單在本月就扣。
    所以這裡不先加總，回傳每天 A＋H 張數與每張退貨單的原單日期，期間加總時再判斷。
    """
    sql = ("select loc_id, to_char(doc_date,'YYYY-MM-DD') d, trans_type tt, "
           "case when trans_type='E' then substr(ref_doc_id, instr(ref_doc_id,'-',1,2)+1, 8) end ref_d, "
           f"count(*) n from posmas where org_id='01' and loc_id in ({LOC_IN}) and {_date_range(d1, d2)} "
           "and trans_type in ('A','H','E') group by loc_id, to_char(doc_date,'YYYY-MM-DD'), trans_type, "
           "case when trans_type='E' then substr(ref_doc_id, instr(ref_doc_id,'-',1,2)+1, 8) end")
    df = F.run_epb(sql)
    sales, returns = {}, []
    for _, r in ([] if df is None else df.iterrows()):
        key = (str(r['LOC_ID']).strip(), str(r['D'])[:10])
        if r['TT'] == 'E':
            ref = str(r['REF_D'])
            ref = f'{ref[:4]}-{ref[4:6]}-{ref[6:8]}' if ref[:8].isdigit() else ''
            returns.append((*key, ref, int(_num(r['N']))))
        else:
            sales[key] = sales.get(key, 0) + int(_num(r['N']))
    return sales, returns


def txn_count(sales, returns, d1, d2):
    lo, hi = d1.isoformat(), d2.isoformat()
    out = {}
    for (loc, d), n in sales.items():
        if lo <= d <= hi:
            out[loc] = out.get(loc, 0) + n
    for loc, d, ref, n in returns:
        if lo <= d <= hi and lo <= ref <= hi:
            out[loc] = out.get(loc, 0) - n
    return out


def periods(day):
    """今日、本月累積（1 日～今日）、上月同期與去年同期（1 日～同一天，該月較短取月底）。"""
    last_month_end = day.replace(day=1) - timedelta(days=1)
    lm_end = last_month_end.replace(day=min(day.day, last_month_end.day))
    ly_start = day.replace(year=day.year - 1, day=1)
    ly_month_end = (ly_start.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    return {'day': (day, day), 'mtd': (day.replace(day=1), day),
            'lm': (last_month_end.replace(day=1), lm_end),
            'ly': (ly_start, ly_start.replace(day=min(day.day, ly_month_end.day)))}


def summarize(daily, d1, d2):
    """把期間內每天的原始數字加總、最後才四捨五入（跟日報同一個順序，避免逐日捨入的尾差）。"""
    totals = {}
    for (loc, d), values in daily.items():
        if d1.isoformat() <= d <= d2.isoformat():
            row = totals.setdefault(loc, {})
            for key, value in values.items():
                row[key] = row.get(key, 0.0) + value
    out = {}
    for loc in LOCS:
        t = totals.get(loc, {})
        g = lambda k: t.get(k, 0.0)
        row = {k: round(v) for k, v in t.items() if not k.endswith(('_p', '_c'))}
        for key in ('gm_total', 'gm_apple', 'gm_tpp', 'acp'):
            row[key] = round(g(f'{key}_p')) - round(g(f'{key}_c'))
        out[loc] = row
    return out


def compute(day):
    """回傳 {loc: {'day','mtd','lm','ly': {...}}} 與 SA Care 對不到價目表的品項。

    分兩段查：上月 1 日～今日（涵蓋今日、本月累積、上月同期）與去年同期那一段。
    """
    ranges = periods(day)
    daily, sales, returns, unmatched = {}, {}, [], []
    for start, end in ((ranges['lm'][0], day), ranges['ly']):
        sacare, missed = query_sacare(start, end)
        unmatched += missed
        for part in (query_sell(start, end), query_margin(start, end), query_acpp_margin(start, end),
                     sacare, query_tradein(start, end)):
            for key, values in part.items():
                daily.setdefault(key, {}).update(values)
        part_sales, part_returns = query_txn(start, end)
        sales.update(part_sales)
        returns += part_returns
    by_period = {}
    for name, (d1, d2) in ranges.items():
        by_period[name] = summarize(daily, d1, d2)
        for loc, n in txn_count(sales, returns, d1, d2).items():
            by_period[name][loc]['txn'] = n
    return {loc: {name: by_period[name][loc] for name in ranges} for loc in LOCS}, unmatched


def intranet_targets(day):
    """內網目標網站當月 28 項（＝日報 target 分頁）：{loc: {'col1': 值, ...}}；連不到就回空的。"""
    try:
        found = P.intranet_targets(day.strftime('%Y%m'))
    except Exception as error:
        print(f'  ⚠️ 連不到內網目標網站，這次只推實績（雲端沿用本月最近一次的目標）：{str(error)[:120]}')
        return {}
    return {loc: {code: value for code, (_, value) in found[f'SA{loc}'].items()}
            for loc in LOCS if found.get(f'SA{loc}')}


def report_failure(day, message):
    """失敗記到雲端的同步紀錄（同步狀態頁看得到）；連雲端都連不上就算了，log 裡還有。"""
    try:
        P.hub_request('/api/agent/report', {'job': f'門市日報 {day}', 'ok': False, 'message': message,
                                           'host': P.host_name()})
    except Exception:
        pass


def push(day, data):
    ids = P.store_ids()
    targets = intranet_targets(day)
    rows = [{'store_id': ids[F.LOC2S[loc]], **values, **({'targets': targets[loc]} if loc in targets else {})}
            for loc, values in data.items()]
    status, body = P.hub_request('/api/epb/daily', {'date': day.isoformat(), 'rows': rows,
                                                    'pushed_by': os.environ.get('USER', '')})
    if status != 200:
        message = f'推送失敗 HTTP {status}：{body.decode("utf-8", "replace")[:300]}'
        report_failure(day, message)
        sys.exit(f'❌ {message}')
    print(f'✅ {day}：{json.loads(body)}')


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    day = date.fromisoformat(args[0]) if args else date.today() - timedelta(days=1)
    if os.environ.get('LAUNCHD'):
        P.use_lib_bridge()                  # launchd 讀不到 ~/Desktop 那份 EPB 橋接
    print(f'[{datetime.now():%Y-%m-%d %H:%M}] 查 EPB 日報 {day} …', flush=True)
    try:
        data, unmatched = compute(day)
    except Exception as error:
        report_failure(day, f'查 EPB 失敗（沒連 VPN？）：{str(error).splitlines()[0][:200]}')
        raise
    F.report_unmatched(f'SA Care {day}', [(F.LOC2S.get(loc, loc), *rest) for loc, *rest in unmatched])
    if '--dry' in sys.argv:
        print(json.dumps(data, ensure_ascii=False, indent=1))
        return
    push(day, data)


if __name__ == '__main__':
    main()
