#!/usr/bin/env python3
"""
ARpedia 日報 - brand_id='496' (BI-B_AR套組)
輸出：ARpedia（昨日/本週/月累/年累） + 人流（昨日/月累/上月）
"""
import subprocess, sys, json
from datetime import date, timedelta
from pathlib import Path

CP   = "/Users/bob.1469/Desktop/north1-weekly-report:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*"
JAVA = "/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java"
CWD  = "/Users/bob.1469/Desktop/north1-weekly-report"
CONFIG = Path("/Users/bob.1469/Desktop/north1-weekly-report/local_config.json")

STORES     = ['士林', '微風', '美麗華', '阿波羅', '高島屋', '羅東']
SHOP_CODES = {'士林': '4', '微風': '5', '美麗華': '24', '阿波羅': '46', '高島屋': '54', '羅東': '57'}
SHOP_IN    = "'004','005','024','046','054','057'"
SHOP_STR   = {'士林':'004','微風':'005','美麗華':'024','阿波羅':'046','高島屋':'054','羅東':'057'}

def epb(sql):
    r = subprocess.run(
        [JAVA, "-Dsun.net.client.defaultReadTimeout=120000",
         "-cp", CP, "EPBReportQuery", sql, "5000"],
        capture_output=True, text=True, cwd=CWD
    )
    lines = [l for l in r.stdout.strip().split('\n') if l.strip()]
    if not lines:
        return []
    hdrs = [h.strip().upper() for h in lines[0].split('\t')]
    return [dict(zip(hdrs, row.split('\t'))) for row in lines[1:]]

def query_units(d_start: date, d_end: date) -> dict:
    ds = f"TO_DATE('{d_start}','yyyy-mm-dd')"
    de = f"TO_DATE('{d_end}','yyyy-mm-dd')"
    rows = epb(
        f"SELECT l.shop_id, "
        f"SUM(CASE WHEN l.trans_type IN ('A','H') THEN l.stk_qty "
        f"WHEN l.trans_type='E' THEN l.stk_qty ELSE 0 END) AS units "
        f"FROM poslinev_bi l "
        f"WHERE l.org_id='01' AND l.shop_id IN ({SHOP_IN}) "
        f"AND l.doc_date>={ds} AND l.doc_date<={de} "
        f"AND l.trans_type IN ('A','E','H') "
        f"AND l.brand_id='496' "
        f"GROUP BY l.shop_id ORDER BY l.shop_id"
    )
    return {str(int(r['SHOP_ID'])): int(float(r.get('UNITS', 0) or 0)) for r in rows}

def query_txn_count(d_start: date, d_end: date, shop_id: str) -> int:
    """羅東無計數器：用 EPB 成交筆數計算人流"""
    ds = f"TO_DATE('{d_start}','yyyy-mm-dd')"
    de = f"TO_DATE('{d_end}','yyyy-mm-dd')"
    rows = epb(
        f"SELECT COUNT(DISTINCT l.doc_id) AS cnt "
        f"FROM poslinev_bi l "
        f"WHERE l.org_id='01' AND l.shop_id='{shop_id}' "
        f"AND l.doc_date>={ds} AND l.doc_date<={de} "
        f"AND l.trans_type IN ('A','H')"
    )
    if rows:
        return int(float(rows[0].get('CNT', 0) or 0))
    return 0

def traffic_formula(txn: int) -> int:
    return round(txn * 0.85 / 0.3)

def fetch_shoppertrak(periods: list[tuple]) -> dict:
    """
    回傳 {store_name: [v1, v2, v3]} 對應三個 period。
    失敗時全部回傳 None。
    """
    import ssl
    ssl._create_default_https_context = ssl._create_unverified_context

    sys.path.insert(0, CWD)
    try:
        import shoppertrak
    except ImportError:
        print("  ⚠️  找不到 shoppertrak 模組，略過人流")
        return {}

    try:
        cfg = json.loads(CONFIG.read_text())
        username = cfg["shoppertrak"]["username"]
        password = cfg["shoppertrak"]["password"]
    except Exception as e:
        print(f"  ⚠️  無法讀取 ShopperTrak 帳密：{e}，略過人流")
        return {}

    # 找出三個 period 的最早到最晚日期，一次抓完
    all_starts = [ds for _, ds, _ in periods]
    all_ends   = [de for _, _, de in periods]
    fetch_start = min(all_starts)
    fetch_end   = max(all_ends)

    st_codes = ['004', '005', '024', '046', '054']  # 羅東無計數器
    try:
        print("  查詢 ShopperTrak 人流...", flush=True)
        daily_map = shoppertrak.fetch_all(
            st_codes, fetch_start, fetch_end, username, password,
            log=lambda m: print(f"  {m}", flush=True)
        )
    except Exception as e:
        print(f"  ⚠️  ShopperTrak 失敗：{e}，略過人流")
        return {}

    name_to_code = SHOP_STR  # {'士林':'004', ...}
    result = {}
    for store in STORES:
        code = name_to_code[store]
        if code == '057':
            continue  # 羅東另算
        days = daily_map.get(code, {})
        vals = []
        for _, ds, de in periods:
            total = 0
            d = ds
            while d <= de:
                total += days.get(d.isoformat(), 0)
                d += timedelta(days=1)
            vals.append(total)
        result[store] = vals

    return result

def main():
    today = date.today()
    yesterday = today - timedelta(days=1)

    # 本週起點：包含 yesterday 的週日（週日～週六制）
    days_since_sunday = (yesterday.weekday() + 1) % 7
    week_start = yesterday - timedelta(days=days_since_sunday)

    month_start = yesterday.replace(day=1)
    year_start  = today.replace(month=1, day=1)

    # 上月同期（1日 ～ 上月的同一天）
    last_month_start = (month_start - timedelta(days=1)).replace(day=1)
    last_month_same  = last_month_start.replace(day=yesterday.day)

    ar_periods = [
        (f"昨日 {yesterday.strftime('%-m/%-d')}",                                          yesterday,   yesterday),
        (f"本週 {week_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",         week_start,  yesterday),
        (f"月累 {month_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",        month_start, yesterday),
        (f"年累 {year_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",         year_start,  yesterday),
    ]

    tr_periods = [
        (f"昨日 {yesterday.strftime('%-m/%-d')}",                                                    yesterday,        yesterday),
        (f"月累 {month_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",                  month_start,      yesterday),
        (f"上月同期 {last_month_start.strftime('%-m/%-d')}~{last_month_same.strftime('%-m/%-d')}",  last_month_start, last_month_same),
    ]

    print(f"\nARpedia、人流 日報 ({today.strftime('%Y-%m-%d')})")
    print("=" * 80)

    # ── ARpedia 銷售 ──
    print("\n【ARpedia 銷售數量】")
    ar_results = []
    for label, ds, de in ar_periods:
        print(f"  查詢 {label}...", flush=True)
        ar_results.append((label, query_units(ds, de)))

    col_w = 12
    header = f"{'門市':<6}" + "".join(f"{label:>{col_w}}" for label, _ in ar_results)
    print("\n" + header)
    print("-" * len(header))
    ar_totals = [0] * len(ar_results)
    for store in STORES:
        sid = SHOP_CODES[store]
        row = f"{store:<6}"
        for i, (_, data) in enumerate(ar_results):
            v = data.get(sid, 0)
            ar_totals[i] += v
            row += f"{v:>{col_w},}"
        print(row)
    print("-" * len(header))
    print(f"{'合計':<6}" + "".join(f"{t:>{col_w},}" for t in ar_totals))

    # ── 人流 ──
    print("\n【人流（來客數）】")
    tr_data = fetch_shoppertrak(tr_periods)

    # 羅東：EPB 成交筆數公式
    roto_vals = []
    for label, ds, de in tr_periods:
        print(f"  查詢羅東成交筆數 {label}...", flush=True)
        txn = query_txn_count(ds, de, '057')
        roto_vals.append(traffic_formula(txn))
    tr_data['羅東'] = roto_vals

    tr_labels = [label for label, _, _ in tr_periods]
    col_w2 = 14
    header2 = f"{'門市':<6}" + "".join(f"{label:>{col_w2}}" for label in tr_labels)
    print("\n" + header2)
    print("-" * len(header2))
    tr_totals = [0] * len(tr_periods)
    for store in STORES:
        vals = tr_data.get(store)
        row = f"{store:<6}"
        for i, v in enumerate(vals or [None]*len(tr_periods)):
            if v is None:
                row += f"{'--':>{col_w2}}"
            else:
                tr_totals[i] += v
                row += f"{v:>{col_w2},}"
        print(row)
    print("-" * len(header2))
    print(f"{'合計':<6}" + "".join(f"{t:>{col_w2},}" for t in tr_totals))
    print()

    # ── 門市日報 本日銷售/休假人數（北一＋北二，來自各店日報信附件）──
    try:
        sys.path.insert(0, str(Path(__file__).parent))
        from daily_headcount import print_headcount
        print("=" * 80)
        print_headcount('北一區')
        print()
    except Exception as e:
        print(f"\n  ⚠️  門市人數統計失敗：{e}")

if __name__ == "__main__":
    main()
