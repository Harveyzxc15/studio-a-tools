#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
目標分配 — EPB 填業績（去年同期 + 上月）
======================================
從 EPB 查各店去年同期 + 上月的實際業績，填進目標表的兩個參考區塊，
供後續「目標分配_四版本.py」用各店結構分配區域目標。

台數/3pp/配件/總銷售＝EPB 精確；SACare 金額 EPB 查不到(代收記0)，
改用「各 SKU 台數 × SAcare 對應價目表」逐品項算，比舊的各店均價估算準。

【每月使用】改 CONFIG 的 TARGET_MONTH + TARGET_FILE，執行：
    python3 目標分配_填業績.py
⚠️ 會覆蓋 OUT_FILE；若之後有手動調整(如加禮券)，重跑會蓋掉，請留意。
======================================
"""
import os, io, subprocess
from datetime import date
import pandas as pd
from openpyxl import load_workbook

# ============== CONFIG（每月改這裡）==============
TARGET_MONTH = "202609"          # 要訂的月份 → 自動推去年同期(202507)、上月(202606)
IMPORT_DIR   = os.path.expanduser(f"~/工具中心/匯入/目標分配/{TARGET_MONTH[:4]}.{TARGET_MONTH[4:]}")  # 依月份自動指向
TARGET_FILE  = f"{IMPORT_DIR}/北ㄧ區門市目標-3.xlsx"          # 原始(未填)目標表
OUT_FILE     = f"{IMPORT_DIR}/北ㄧ區門市目標-3_填業績.xlsx"   # 填好業績的目標表

# SACare 金額在 EPB 是代收、所有金額欄都 0，只查得到台數，所以價格要自己帶。
# 價目表＝公司給的「SAcare對應價目表」(存代→價格)，已驗證與品名內建的 ($X,XXX) 100% 一致。
# ⚠️ 新機檔期會出新 SKU，跑完若印出「未在價目表」清單，跟公司要新版價目表換掉這個檔。
SACARE_PRICE_FILE = os.path.expanduser("~/工具中心/lib/SAcare對應價目表.xlsx")
# ==================================================

CWD = os.path.expanduser("~/Desktop/daily-tracker")   # EPBReportQuery 所在
CP = f"{CWD}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*"
CUST = ("('ST00237177','SA03189128','SA03147688','SA03125724','SA03132099','SA03145281',"
        "'SA03129217','SA03130573','SA03175171','SA03147048','SA03184320')")
LOC2S = {'004': '士林', '005': '微風', '024': '美麗華', '046': '阿波羅', '054': '高島屋', '057': '羅東',
         # 北二七店：本腳本產的目標表用不到，是給 目標分配_推雲端.py 一次查 13 店用的。
         '009': '永和', '025': '板橋誠品', '050': '新西門', '055': '花蓮',
         '063': '板橋遠百', '064': '新莊宏匯', '068': '新店裕隆城'}
NORTH1 = ('004', '005', '024', '046', '054', '057')
NORTH2 = ('009', '025', '050', '055', '063', '064', '068')
STORES = ['士林', '微風', '美麗華', '阿波羅', '高島屋', '羅東']


def run_epb(sql):
    r = subprocess.run(['java', '-cp', CP, 'EPBReportQuery', sql, '99999'],
                       capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=CWD)
    return pd.read_csv(io.StringIO(r.stdout), sep='\t', dtype=str) if r.stdout.strip() else None


def base_where(d1, d2, locs=NORTH1, trans="l.trans_type not in ('G','J')"):
    """共用的期間/門市/單別/客戶排除條件。locs 預設北一六店，推雲端會傳 13 店。"""
    loc_list = ','.join(f"'{loc}'" for loc in locs)
    return (f"l.org_id='01' and l.loc_id in ({loc_list}) and {trans} "
         f"and l.app_code in ('RNCN','DNN','POSN') and (l.cust_id is null or l.cust_id not in {CUST}) "
         f"and l.doc_date>=to_date('{d1}','yyyy-mm-dd') and l.doc_date<to_date('{d2}','yyyy-mm-dd')")


def query(d1, d2, locs=NORTH1):
    """查一期間各店：7品類台數 + 3pp(cat3=3003) + 配件(cat3=3002) + 總銷售 + SACare台數。

    口徑照門市日報（2026/9 用 7 店日報驗證）：
      金額＝A 銷售＋E 銷退＋G 訂金＋J 退訂（訂金收到當下就算業績＝日報「零售訂金營業額(c)」），
          不算 H 訂單出貨（日報列成 (b)、不進總營業額，錢已經在 G 訂金算過）
      台數＝A＋E＋H（G／J 沒有台數，跟日報台數一致）
    """
    B = base_where(d1, d2, locs, trans="l.trans_type in ('A','E','G','H','J')")
    qty = "l.trans_type not in ('G','J')"      # 台數
    amt = "l.trans_type<>'H'"                  # 金額
    sql = (f"select l.loc_id, "
           f"sum(case when {qty} and l.cat4_id in(4001,4002) and l.cat5_id<>5002 then l.stk_qty else 0 end) cpu, "
           f"sum(case when {qty} and l.cat4_id in(4001,4002) and l.cat5_id=5002 then l.stk_qty else 0 end) mini, "
           f"sum(case when {qty} and l.cat4_id in(4005,4006,4041) and l.cat5_id in(5099,5507,5467,5466,5559,5476) then l.stk_qty else 0 end) ipad, "
           f"sum(case when {qty} and l.cat4_id in(4005,4006,4041) and l.cat5_id=5107 then l.stk_qty else 0 end) ipadmini, "
           f"sum(case when {qty} and l.cat4_id in(4005,4006,4041) and l.cat5_id in(5560,5583) then l.stk_qty else 0 end) ipadpro, "
           f"sum(case when {qty} and l.cat4_id=4004 then l.stk_qty else 0 end) iphone, "
           f"sum(case when {qty} and l.cat4_id=4038 then l.stk_qty else 0 end) watch, "
           f"round(sum(case when {amt} and l.cat3_id='3003' then l.line_total_aftdisc else 0 end)) tpp, "
           f"round(sum(case when {amt} and l.cat3_id='3002' then l.line_total_aftdisc else 0 end)) acc, "
           f"round(sum(case when {amt} then l.line_total_aftdisc else 0 end)) totsale, "
           f"sum(case when {qty} and upper(l.name) like 'S.A CARE%' then l.stk_qty else 0 end) sacqty "
           f"from enqsales_bi l where {B} group by l.loc_id order by l.loc_id")
    df = run_epb(sql)
    keys = ['cpu', 'mini', 'ipad', 'ipadmini', 'ipadpro', 'iphone', 'watch', 'tpp', 'acc', 'totsale', 'sacqty']
    return {LOC2S[r['LOC_ID']]: {k: int(float(r[k.upper()])) for k in keys} for _, r in df.iterrows()}


def query_sacare(d1, d2, base):
    """各店 SACare 金額＝Σ(該 SKU 台數 × 價目表價格)。

    EPB 的 SACare 是代收，金額欄全是 0，只能靠台數×價格還原。
    對不到價目表的幾乎都是「維修費」品項(非保固方案)，本來就不該計入金額，記 0 並列出來給人看。
    """
    prices = pd.read_excel(SACARE_PRICE_FILE, header=0)
    prices.columns = ['stk_id', 'name', 'price']
    prices['stk_id'] = prices['stk_id'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    lookup = prices.dropna(subset=['stk_id']).drop_duplicates('stk_id').set_index('stk_id')['price']

    sql = (f"select l.loc_id, l.stk_id, max(l.name) nm, sum(l.stk_qty) qty from enqsales_bi l "
           f"where {base} and upper(l.name) like 'S.A CARE%' group by l.loc_id, l.stk_id")
    df = run_epb(sql)
    if df is None:
        return {}, []

    df['qty'] = pd.to_numeric(df['QTY'])
    df['price'] = df['STK_ID'].str.strip().map(lookup)
    unmatched = [(LOC2S[r['LOC_ID']], r['STK_ID'], r['NM'], r['qty'])
                 for _, r in df[df['price'].isna() & (df['qty'] != 0)].iterrows()]
    df['amount'] = df['qty'] * df['price'].fillna(0)
    totals = df.groupby('LOC_ID')['amount'].sum()
    return {LOC2S[loc]: round(amount) for loc, amount in totals.items()}, unmatched


def month_ranges(ym):
    """由 TARGET_MONTH 推去年同期、上月的日期範圍。"""
    y, m = int(ym[:4]), int(ym[4:])
    def rng(yy, mm):
        d1 = date(yy, mm, 1)
        d2 = date(yy + 1, 1, 1) if mm == 12 else date(yy, mm + 1, 1)
        return d1.isoformat(), d2.isoformat()
    ly = rng(y - 1, m)                                  # 去年同期
    pm_y, pm_m = (y, m - 1) if m > 1 else (y - 1, 12)   # 上月
    m6 = rng(pm_y, pm_m)
    return ly, m6, f"{y-1}/{m}", f"{pm_y}/{pm_m}"


def fill_block(ws, base_row, data, sacare):
    for i, s in enumerate(STORES):
        r, d = base_row + i, data[s]
        sac = sacare.get(s, 0)
        ws.cell(r, 2, d['totsale'] + sac)   # 總業績(含SACare)
        ws.cell(r, 3, d['tpp'])             # 3pp
        ws.cell(r, 4, sac)                  # SACare(估算)
        for j, cat in enumerate(['cpu', 'mini', 'ipad', 'ipadmini', 'ipadpro', 'iphone', 'watch']):
            ws.cell(r, 7 + j, d[cat])       # 台數
        ws.cell(r, 14, d['acc'])           # Apple配件


def report_unmatched(label, unmatched):
    """對不到價目表的 SKU 一定要講出來，不然金額會靜默少算。"""
    if not unmatched:
        return
    total = sum(qty for *_, qty in unmatched)
    print(f"\n⚠️ {label} 有 {len(unmatched)} 個 SKU（共 {total:.0f} 台）不在價目表，金額以 0 計：")
    for store, stk_id, name, qty in unmatched:
        print(f"     {store:<5}{stk_id:<12}{str(name)[:40]:<42}{qty:>5.0f}")
    print("     維修費類本來就不該計入金額；若出現保固方案，請跟公司要新版價目表")


def main():
    ly_rng, m6_rng, ly_lbl, m6_lbl = month_ranges(TARGET_MONTH)
    print(f"去年同期={ly_lbl}  上月={m6_lbl}")
    ly = query(*ly_rng)
    m6 = query(*m6_rng)
    ly_sac, ly_miss = query_sacare(*ly_rng, base_where(*ly_rng))
    m6_sac, m6_miss = query_sacare(*m6_rng, base_where(*m6_rng))
    wb = load_workbook(TARGET_FILE); ws = wb.active
    fill_block(ws, 20, ly, ly_sac)   # 去年同期 → row20-25
    fill_block(ws, 30, m6, m6_sac)   # 上月     → row30-35
    wb.save(OUT_FILE)
    print(f"✅ 已填: {OUT_FILE}")
    print("   台數/3pp/配件/總業績=EPB精確；SACare 金額=各SKU台數×價目表")
    # 快速核對
    print(f"\n{'門市':<6}{'去年iPhone':>10}{'上月iPhone':>10}{'上月總業績':>13}"
          f"{'去年SACare':>12}{'上月SACare':>12}")
    for s in STORES:
        print(f"{s:<6}{ly[s]['iphone']:>10}{m6[s]['iphone']:>10}"
              f"{m6[s]['totsale']+m6_sac.get(s, 0):>13,}{ly_sac.get(s, 0):>12,}{m6_sac.get(s, 0):>12,}")
    report_unmatched(f"去年同期({ly_lbl})", ly_miss)
    report_unmatched(f"上月({m6_lbl})", m6_miss)


if __name__ == "__main__":
    main()
