"""匯出檔模式：不連 EPB，所有 EPB 查詢改讀 EPBrowser 手動匯出檔（~/工具中心/匯入/EPB匯出/*.xlsx）。

2026-11-17 使用者不想讓 EPB 端看到程式查詢（auto-memory epb_gui_export_source）。enable() 之後：
- 任何子程序只要帶 EPBReportQuery／EPBQueryGuard／EPBDocPdf（Java 橋接）就直接報錯 → 不會偷偷連 EPB。
- warranty_report_all.epb(sql) 改在 SQLite 的 poslinev_bi（匯出檔建的）上跑同一句 SQL。
- 週報產生器引擎 load_from_epb／load_txn_from_epb 改用 export_loader。
- 查詢期間超過匯出檔涵蓋的日期 → 報錯（不拿缺資料的期間照算）。
"""
import datetime as dt
import os
import re
import sqlite3
import subprocess
import threading
from pathlib import Path

import pandas as pd

import export_loader as X
from daily_facts import STORES

EXPORT_DIR = Path(os.environ.get('EPB_EXPORT_DIR') or Path.home() / '工具中心' / '匯入' / 'EPB匯出')
BLOCK = ('EPBReportQuery', 'EPBQueryGuard', 'EPBDocPdf', 'EPBrowser')
TT = {'銷售': 'A', '銷退': 'E', '訂金': 'G', '尾款': 'H', '退訂': 'J'}
ALL = [c for st in STORES.values() for c in st]          # 13 店門市代碼
_state = {}
_lock = threading.Lock()                               # 保固圖用多條執行緒同時查：建表與查詢都要排隊


class EPBBlocked(RuntimeError):
    pass


def scan():
    """(可用的檔, {不能用的檔名: 原因})。傳到一半、不是匯出檔、缺欄位的檔跳過，不讓整天失敗。"""
    good, bad = [], {}
    for p in sorted(EXPORT_DIR.glob('*.xlsx')):
        if p.name.startswith(('~$', '.')):
            continue
        why = X.check_file(p)
        if why:
            bad[p.name] = why
        else:
            good.append(str(p))
    return good, bad


def files():
    if 'files' not in _state:
        _state['files'], _state['bad'] = scan()
    return _state['files']


def refresh():
    """重新掃資料夾（等檔時用）；SQLite 也要重建。"""
    for k in ('files', 'bad', 'db'):
        _state.pop(k, None)


def check_covered(start, end, shops=None):
    """shops（門市代碼）在 [start, end] 每一天都有匯出檔涵蓋；有缺就報錯（不拿缺資料的期間照算）。"""
    shops = shops or ALL
    lack = X.missing_days(files(), start, end, shops)
    if lack:
        days = sorted({d for v in lack.values() for d in v})
        span = f'{days[0]:%m/%d}' + (f'～{days[-1]:%m/%d}（{len(days)} 天）' if len(days) > 1 else '')
        raise RuntimeError(f'匯出檔沒涵蓋 {span}：{"、".join(sorted(lack))}')


def _guard():
    orig = subprocess.Popen

    class Guarded(orig):
        def __init__(self, args, *a, **k):
            s = ' '.join(map(str, args)) if isinstance(args, (list, tuple)) else str(args)
            if any(b in s for b in BLOCK):
                raise EPBBlocked('匯出檔模式禁止連 EPB：' + s[:150])
            super().__init__(args, *a, **k)

    subprocess.Popen = Guarded          # subprocess.run／check_output 內部都走 subprocess.Popen


def _db():
    if 'db' in _state:
        return _state['db']
    raw, d = X.select_raw(files(), '2000-01-01', '2100-12-31')
    txt = lambda c: raw[c].str.strip().replace('', None)
    num = lambda c: pd.to_numeric(raw[c], errors='coerce').fillna(0.0)
    df = pd.DataFrame({
        'org_id': '01', 'shop_id': raw['門市代碼'], 'doc_id': raw['單據代碼'].str.strip(),
        'doc_date': d.dt.strftime('%Y-%m-%d'), 'trans_type': raw['交易類型'].str.strip().map(TT),
        'emp_id1': txt('員工代碼'), 'stk_id': txt('存貨代碼'), 'name': raw['名稱'].str.strip(),
        'stk_qty': num('數量'), 'line_total_net': num('淨銷售金額(未稅)'), 'line_tax': num('淨銷售稅額'),
        'cost_price': num('單位成本'), 'trn_cost_price': num('單位成本'),
        'cat1_id': txt('類別1代碼'), 'cat2_id': txt('類別2代碼'), 'cat3_id': txt('類別3代碼'),
        'cat4_id': txt('類別4代碼'), 'cat5_id': txt('類別5代碼'), 'cat6_id': txt('類別6代碼'),
        'brand_id': txt('品牌代碼'),
    })
    con = sqlite3.connect(':memory:', check_same_thread=False)
    con.execute('PRAGMA case_sensitive_like = ON')          # Oracle 的 LIKE 分大小寫
    df.to_sql('poslinev_bi', con, index=False)
    con.execute('CREATE INDEX i1 ON poslinev_bi (doc_date, shop_id)')
    con.execute('CREATE INDEX i2 ON poslinev_bi (doc_id)')
    _state['db'] = con
    return con


_TO_DATE = re.compile(r"TO_DATE\(\s*'(\d{4}-\d{2}-\d{2})'\s*,\s*'yyyy-mm-dd'\s*\)", re.I)
_TO_CHAR = re.compile(r"to_char\(\s*([\w.]+)\s*,\s*'YYYY-MM-DD'\s*\)", re.I)
_SHOPS = re.compile(r"shop_id\s+IN\s*\(([^)]*)\)", re.I)


def epb(sql, max_rows=10000):
    """warranty_report_all.epb 的替身：同一句 Oracle SQL 在 SQLite 上跑，回傳格式一樣（大寫欄名、字串值）。"""
    dates = _TO_DATE.findall(sql)
    if dates:
        m = _SHOPS.search(sql)
        shops = [s.strip(" '") .zfill(3) for s in m.group(1).split(',')] if m else None
        hi = max(dates)
        if 'DOC_DATE<TO_DATE' in sql.replace(' ', '').upper():        # 「< 迄日+1」的寫法
            hi = (dt.date.fromisoformat(hi) - dt.timedelta(days=1)).isoformat()
        with _lock:                                    # files()／missing_days 也共用 _state
            check_covered(min(dates), hi, shops)
    q = _TO_CHAR.sub(r'substr(\1,1,10)', _TO_DATE.sub(r"'\1'", sql))
    with _lock:
        cur = _db().execute(q)
        cols = [c[0].split('.')[-1].upper() for c in cur.description]
        rows = cur.fetchmany(max_rows)
    return [{k: '' if v is None else str(v) for k, v in zip(cols, r)} for r in rows]


def _patch_engine(eng):
    def load_from_epb(start_date, end_date, store_codes=None, org_id='01'):
        check_covered(start_date, end_date, store_codes)
        return X.load(files(), start_date, end_date, store_codes)

    def load_txn_from_epb(start_date, end_date, store_codes, org_id='01'):
        return load_from_epb(start_date, end_date, store_codes).attrs['txn']

    eng.load_from_epb, eng.load_txn_from_epb = load_from_epb, load_txn_from_epb
    return eng


def enable():
    """開啟匯出檔模式。要在 import warranty_report_all／week_facts 的程式開始查詢前呼叫。"""
    if _state.get('on'):
        return
    _guard()
    import warranty_report_all as W
    W.epb = epb
    import week_facts as WF
    orig = WF.load_engine
    WF.load_engine = lambda region: _patch_engine(orig(region))
    _state['on'] = True
