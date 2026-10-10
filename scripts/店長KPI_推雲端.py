#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
店長 KPI — 用 EPB 算 13 店的店長 KPI 原始數字，推到 sa-north-hub
======================================
口徑照「北一區-YYYY.MM月份店長KPI統計.xlsx」的填表規則（記憶 store_manager_kpi_file）：
  D  CPU 台數   ＝cat4 4001/4002 淨台數，含認證機
  K  Mac 保固   ＝SACare cat6 6533/6546 ＋ 認證機保固 99902051 ＋ ACPP cat3 3032 品名含 Mac
  R  iOS 保固   ＝分母 cat4 4004/4005/4006/4041 排除認證機 cat2 2029；
                  分子 SACare cat6 6534/6535 ＋ ACPP cat3 3032 品名含 iphone/ipad
  Y  3PP 搭售率 ＝dss_rebuild_from_epb 重建 DSS 搭售明細，Σ五類配件 ÷ Σ五類搭售量
  AF 交辦事項   ＝雲端頁面設定的條件（cat4＋品名關鍵字）淨台數
淨數量一律 trans_type A/E/H 直接加總 stk_qty（銷退本身就是負的）。

跨店教育價（羅東／永和／花蓮／板橋誠品的人在別店結教育單）在雲端做：
這裡只推「含教育單 SKU 的單據」依 結帳店×員工 的 Mac 台數與保固數，
雲端用人員名單的主掛門市決定要不要歸回原店。

【用法】
    python3 店長KPI_推雲端.py 202609          # 查 2026/09 整月並推上去
    python3 店長KPI_推雲端.py 202609 --dry    # 只印不推
    （平常由 目標分配_推雲端.py --poll 依雲端要求自動呼叫：每月 1 日推上個月、頁面按「刷新」推指定月）

【Token】同 目標分配_推雲端.py：EPB_PUSH_TOKEN 或 ~/.config/studioa/epb_push_token
======================================
"""
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dss_rebuild_from_epb as D

SHOPS = D.REGIONS['01'] + D.REGIONS['02']
SHOP_IN = ','.join(f"'{s}'" for s in SHOPS)
EDU_SKUS = "'99200202','99200203'"

MAC = "cat4_id IN ('4001','4002')"
MAC_CARE = ("(cat6_id IN ('6533','6546') OR stk_id='99902051'"
            " OR (cat3_id='3032' AND name LIKE '%Mac%'))")
IOS = "(cat4_id IN ('4004','4005','4006','4041') AND (cat2_id<>'2029' OR cat2_id IS NULL))"
IOS_CARE = ("(cat6_id IN ('6534','6535') OR (cat3_id='3032'"
            " AND (LOWER(name) LIKE '%iphone%' OR LOWER(name) LIKE '%ipad%')))")


def month_range(ym):
    y, m = int(ym[:4]), int(ym[4:6])
    start = date(y, m, 1)
    end = date(y + (m == 12), m % 12 + 1, 1)
    return start, end


def where(ym):
    start, end = month_range(ym)
    return (f"org_id='01' AND shop_id IN ({SHOP_IN}) AND trans_type IN ('A','E','H')"
            f" AND doc_date >= to_date('{start}','YYYY-MM-DD') AND doc_date < to_date('{end}','YYYY-MM-DD')")


def sql_text(value):
    return str(value).replace("'", "''")


def mission_sql(mission):
    """交辦事項條件：cat4 清單（逗號分隔）＋品名關鍵字（不分大小寫）。兩個都空就不算。"""
    parts = []
    cat4 = [c.strip() for c in str(mission.get('cat4') or '').split(',') if c.strip().isdigit()]
    if cat4:
        parts.append('cat4_id IN (' + ','.join(f"'{c}'" for c in cat4) + ')')
    keyword = str(mission.get('keyword') or '').strip()
    if keyword:
        parts.append(f"LOWER(name) LIKE '%{sql_text(keyword.lower())}%'")
    return ' AND '.join(parts) or None


def num(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def query(ym, mission=None):
    w = where(ym)
    m_sql = mission_sql(mission or {})
    mission_col = f", SUM(CASE WHEN {m_sql} THEN stk_qty ELSE 0 END) mission" if m_sql else ''
    agg = D.epb(
        f"SELECT shop_id, SUM(CASE WHEN {MAC} THEN stk_qty ELSE 0 END) mac,"
        f" SUM(CASE WHEN {MAC_CARE} THEN stk_qty ELSE 0 END) mac_care,"
        f" SUM(CASE WHEN {IOS} THEN stk_qty ELSE 0 END) ios,"
        f" SUM(CASE WHEN {IOS_CARE} THEN stk_qty ELSE 0 END) ios_care{mission_col}"
        f" FROM poslinev_bi WHERE {w} GROUP BY shop_id")
    stores = {s: {'mac': 0, 'mac_care': 0, 'ios': 0, 'ios_care': 0, 'mission': None if not m_sql else 0,
                  'tpp_host': 0, 'tpp_acc': 0} for s in SHOPS}
    for r in agg.itertuples():
        row = stores[r.SHOP_ID]
        row.update(mac=num(r.MAC), mac_care=num(r.MAC_CARE), ios=num(r.IOS), ios_care=num(r.IOS_CARE))
        if m_sql:
            row['mission'] = num(r.MISSION)

    # 教育單：含教育單 SKU 的單據，依 結帳店×員工 加總 Mac 與 Mac 保固（雲端決定要不要歸回主掛門市）
    edu = D.epb(
        f"SELECT shop_id, emp_id1, SUM(CASE WHEN {MAC} THEN stk_qty ELSE 0 END) mac,"
        f" SUM(CASE WHEN {MAC_CARE} THEN stk_qty ELSE 0 END) mac_care"
        f" FROM poslinev_bi WHERE {w} AND doc_id IN (SELECT doc_id FROM poslinev_bi WHERE {w}"
        f" AND stk_id IN ({EDU_SKUS})) GROUP BY shop_id, emp_id1")
    edu_rows = [{'shop': r.SHOP_ID, 'emp_id': r.EMP_ID1, 'mac': num(r.MAC), 'mac_care': num(r.MAC_CARE)}
                for r in edu.itertuples() if num(r.MAC) or num(r.MAC_CARE)]

    # 3PP：重建 DSS 搭售明細，搭售量＝「X搭售」主機數量、配件＝「X配件」數量
    start, end = month_range(ym)
    last = date.fromordinal(end.toordinal() - 1)
    lines, occ = D.pull(str(start), str(last), SHOPS)
    tagged = D.label(lines, occ)
    if len(tagged):
        qty = tagged.STK_QTY.astype(float)
        host = tagged.type_name.str.endswith('搭售') & ~tagged.type_name.str.contains('零')
        acc = tagged.type_name.str.endswith('配件')
        for shop, value in qty[host].groupby(tagged.SHOP_ID[host]).sum().items():
            stores[shop]['tpp_host'] = int(value)
        for shop, value in qty[acc].groupby(tagged.SHOP_ID[acc]).sum().items():
            stores[shop]['tpp_acc'] = int(value)
    return {'ym': ym, 'stores': stores, 'edu': edu_rows, 'mission': mission or None,
            'data_end': str(min(date.today(), end))}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if not args or len(args[0]) != 6 or not args[0].isdigit():
        sys.exit('用法：python3 店長KPI_推雲端.py YYYYMM [--dry]')
    ym = args[0]
    import 目標分配_推雲端 as P
    mission = None
    status, body = P.hub_request(f'/api/manager-kpi/mission?ym={ym}')
    if status == 200:
        mission = json.loads(body).get('mission')
    data = query(ym, mission)
    data['pushed_by'] = P.host_name()
    if '--dry' in sys.argv:
        print(json.dumps(data, ensure_ascii=False, indent=1))
        return
    status, body = P.hub_request('/api/manager-kpi/push', data)
    if status != 200:
        sys.exit(f'❌ 推送失敗 HTTP {status}：{body[:300]!r}')
    print(f'✅ 店長 KPI {ym} 已推上 sa-north-hub')


if __name__ == '__main__':
    main()
