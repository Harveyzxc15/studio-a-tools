#!/usr/bin/env python3
"""用 EPB 直接重建 DSS 3PP 搭售明細（免登入 DSS，可排程）

    python3 dss_rebuild_from_epb.py 2026-08-02 2026-08-08 --csv 明細.csv

2026-08-14 定案。對 DSS 黃金樣本驗證（兩區各四週，共 4,401 台主機）：
    北一區  筆數全對、店別+人員彙總 100%、逐筆貼標 99.5~100%
    北二區  筆數全對、彙總 98.2~99.4%、逐筆貼標 99.6~99.9%

⚠️ 這是「從 DSS 輸出反推」的規則，不是 DSS 的原始邏輯。已驗證 poslinev_bi
   並未儲存「哪個配件配哪台主機」，該配對是 DSS 自行計算，因此無法保證完全一致。
   殘留差異：北二區 4 台（客戶同日多單的配對情境），其餘皆為不影響加總的
   tie-break（同品類多台主機時標在哪一台）。
👉 對外要與全台灣對齊的數字（獎金／KPI／跨區比較）請用 DSS 匯出；
   本工具用於 DSS 做不到的分析與快速預覽。用 dss_regression_test.py 守住口徑。
"""
import argparse
import subprocess
from pathlib import Path

import pandas as pd

JAVA = "/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java"
# EPBReportQuery.class 放在 ~/工具中心/lib，不要指向 ~/Desktop —— Desktop 受 macOS
# 隱私權限保護，一旦權限被收回整條工具鏈就會「找不到主要類別」（2026-08-14 踩過）
LIB = str(Path.home() / "工具中心" / "lib")
CP = f"{LIB}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*"
EPB_ROOT = "/Library/EPBrowser"   # 這版 EPBReportQuery 的第一個參數（讀 EPB/Setting.xml）
REGIONS = {
    "01": ("004", "005", "024", "046", "054", "057"),                  # 北一區
    "02": ("009", "025", "050", "055", "063", "064", "068"),           # 北二區
}
SHOPS = REGIONS["01"]

# --- DSS 口徑（逆向自黃金樣本）---
HOST_CAT4 = {"4001": "CPU", "4002": "CPU", "4004": "iPhone",
             "4005": "iPad", "4006": "iPad", "4041": "iPad",
             "4038": "Watch", "4014": "Airpods"}
HOST_CAT3 = {"3001", "3002", "3033", "3035"}   # 主機 / APL週邊(AirPods) / 認證機 / 福利機
AIRPODS_CAT5 = {"5087", "5463", "5599"}        # 排除 5035 EarPods 有線耳機；5599=AirPods 第5代（2026-10 新增）
ACC_CAT3 = "3003"                              # 3PP 週邊配件
TRANS_TYPES = ("A", "E", "H")
EXCLUDE_OCCUPATION = ("88", "99")              # 企業公司戶 / 通訊業
# 配件歸屬：籃內「單行金額(絕對值)最高」的那台主機所屬品類。
# 2026-08-14 用四週 29 張多主機單據驗證 29/29；固定品類優先序最好只有 28/29。


def epb(sql, max_rows=200000):
    out = subprocess.run(
        [JAVA, "-Dsun.net.client.defaultConnectTimeout=8000",
         "-Dsun.net.client.defaultReadTimeout=600000",
         "-cp", CP, "EPBReportQuery", EPB_ROOT, sql, str(max_rows)],
        cwd=LIB,
        capture_output=True, text=True, timeout=900)
    both = out.stdout + out.stderr      # 例外寫 stdout、MAX_ROWS_REACHED 寫 stderr
    if "Exception" in both[:3000] or not out.stdout.strip():
        raise SystemExit(f"❌ EPB 查詢失敗：\n{both[:600]}")
    # 命中筆數上限代表資料被截斷 —— 絕不能安靜吞掉，否則報表會少算
    if "MAX_ROWS_REACHED" in both:
        raise SystemExit(f"❌ EPB 回傳達上限 {max_rows} 筆，資料已被截斷。請縮小日期區間或調高 max_rows。")
    from io import StringIO
    df = pd.read_csv(StringIO(out.stdout), sep="\t", dtype=str)
    return df[df.iloc[:, 0].notna()]


def pull(start, end, shops=SHOPS):
    shop_in = ",".join(f"'{s}'" for s in shops)
    rng = (f"doc_date BETWEEN to_date('{start}','YYYY-MM-DD') "
           f"AND to_date('{end}','YYYY-MM-DD')")
    lines = epb(
        "SELECT doc_id, to_char(doc_date,'YYYY-MM-DD') doc_date, shop_id, store_id,"
        " emp_id1, stk_id, name, stk_qty, cat3_id, cat4_id, cat5_id, trans_type, vip_id,"
        " line_no, line_total"
        f" FROM poslinev_bi WHERE org_id='01' AND shop_id IN ({shop_in}) AND {rng}"
        f" AND cat3_id IN ('3001','3002','3003','3033','3035')"
        " ORDER BY doc_id, line_no")   # 固定順序，金額平手時結果才可重現
    occ = epb(
        "SELECT DISTINCT v.vip_id, v.occupation_id FROM posvip v WHERE v.vip_id IN"
        f" (SELECT DISTINCT l.vip_id FROM poslinev_bi l WHERE l.org_id='01'"
        f" AND l.shop_id IN ({shop_in}) AND l.{rng})", 50000)
    return lines, occ


def label(e, occ):
    drop = set(occ[occ.OCCUPATION_ID.isin(EXCLUDE_OCCUPATION)].VIP_ID)
    is_host = (e.CAT3_ID.isin(HOST_CAT3) & e.CAT4_ID.isin(HOST_CAT4)
               & ~((e.CAT4_ID == "4014") & ~e.CAT5_ID.isin(AIRPODS_CAT5)))
    e = e.assign(role=None, cat=e.CAT4_ID.map(HOST_CAT4).where(is_host))
    e.loc[is_host, "role"] = "host"
    e.loc[e.CAT3_ID == ACC_CAT3, "role"] = "acc"
    e = e[e.role.notna() & e.TRANS_TYPE.isin(TRANS_TYPES) & ~e.VIP_ID.isin(drop)].copy()

    # 用 LINE_TOTAL（未扣折扣的原價）而非 LINE_TOTAL_NET：
    # H 類分期交貨的 NET 是 0，兩台主機都 0 就會平手、變成看資料列順序。
    hosts = e[e.role == "host"].assign(
        amt=pd.to_numeric(e.LINE_TOTAL, errors="coerce").fillna(0).abs())

    def winner(grp):
        """籃內單行金額最高的主機，其品類即為整籃配件的歸屬。"""
        return grp.loc[grp.amt.idxmax(), "cat"]

    by_doc = {k: winner(v) for k, v in hosts.groupby("DOC_ID")}
    # 跨單搭售限「同一家店」—— 2026-08-14 北二區驗證：同客戶同日跨店 DSS 不連結
    by_vip = {k: winner(v) for k, v in hosts.groupby(["VIP_ID", "DOC_DATE", "SHOP_ID"])}

    def basket(r):
        return by_doc.get(r.DOC_ID) or by_vip.get((r.VIP_ID, r.DOC_DATE, r.SHOP_ID))

    def bkey(r):
        if r.DOC_ID in by_doc:
            return ("D", r.DOC_ID)
        return ("V", r.VIP_ID, r.DOC_DATE, r.SHOP_ID) if pd.notna(r.VIP_ID) else ("D", r.DOC_ID)

    attached = {(bkey(r), basket(r)) for r in e.itertuples()
                if r.role == "acc" and basket(r)}
    # 該 (客戶, 日期, 品類) 是否已經有「配件與主機同單」的搭售成立
    doc_level = {(r.VIP_ID, r.DOC_DATE, r.SHOP_ID, basket(r)) for r in e.itertuples()
                 if r.role == "acc" and basket(r) and bkey(r)[0] == "D" and pd.notna(r.VIP_ID)}

    e["amt"] = pd.to_numeric(e.LINE_TOTAL, errors="coerce").fillna(0).abs()

    def slot_of(r):
        """這台主機競逐哪個搭售名額；不符資格回 None。"""
        k = (bkey(r), r.cat)
        vk = (("V", r.VIP_ID, r.DOC_DATE, r.SHOP_ID), r.cat)
        if k in attached:
            return k
        # 跨單落單的配件只有在「該客戶當天該品類還沒有同單搭售」時，
        # 才會把落單主機升級成搭售（2026-08-14 驗證）。
        if (pd.notna(r.VIP_ID) and vk in attached
                and (r.VIP_ID, r.DOC_DATE, r.SHOP_ID, r.cat) not in doc_level):
            return vk
        return None

    # 每個名額只有一台能標搭售，給「單價最高」的那台 —— 與配件歸屬同一原則。
    # 2026-08-14 驗證：價格有高低差的 8 組 tie-break，DSS 全部選較貴那台。
    # ⚠️ 曾試過「價格相同時選 STK_ID 較小者」（4/4 命中），全量驗證反而讓
    #    對不上的從 16 筆增為 28 筆 —— 樣本太小的過擬合，已撤回。價格相同時
    #    維持資料列順序即可，反正不影響任何加總。
    best = {}
    for r in e.itertuples():
        if r.role != "host":
            continue
        s = slot_of(r)
        if s is not None and (s not in best or r.amt > best[s][0]):
            best[s] = (r.amt, r.Index)
    winner = {s: idx for s, (_, idx) in best.items()}

    rows = []
    for r in e.itertuples():
        c = basket(r)
        if r.role == "acc":
            if not c:
                continue
            tag = c + "配件"
        else:
            s = slot_of(r)
            tag = r.cat + ("搭售" if s is not None and winner.get(s) == r.Index
                           else "零搭售")
        rows.append({"DOC_ID": r.DOC_ID, "DOC_DATE": r.DOC_DATE, "SHOP_ID": r.SHOP_ID,
                     "EMP_ID1": r.EMP_ID1, "VIP_ID": r.VIP_ID, "STK_ID": r.STK_ID,
                     "NAME": r.NAME, "STK_QTY": r.STK_QTY, "TRANS_TYPE": r.TRANS_TYPE,
                     "CAT3_ID": r.CAT3_ID, "CAT4_ID": r.CAT4_ID, "CAT5_ID": r.CAT5_ID,
                     "LINE_TOTAL": r.LINE_TOTAL, "type_name": tag})
    return pd.DataFrame(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("start")
    ap.add_argument("end")
    ap.add_argument("--csv", type=Path)
    ap.add_argument("--dept", default="01", choices=sorted(REGIONS),
                    help="01=北一區 02=北二區")
    a = ap.parse_args()

    lines, occ = pull(a.start, a.end, REGIONS[a.dept])
    df = label(lines, occ)
    print(f"{a.start} ~ {a.end}：重建 {len(df)} 筆")
    if a.csv:
        df.to_csv(a.csv, index=False, encoding="utf-8-sig")
        print(f"✅ {a.csv}")
    else:
        print(df.type_name.value_counts().to_string())


if __name__ == "__main__":
    main()
