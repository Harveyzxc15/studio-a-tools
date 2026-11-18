"""EPBrowser「BISHOP & 銷售分析作業」手動匯出 xlsx → 週報產生器引擎的 DataFrame（取代 load_from_epb）。

2026-11-17 起使用者停掉 EPB 直查，改手動匯出（auto-memory epb_gui_export_source）。
金額口徑同 load_from_epb：NET＝淨銷售金額(未稅)＋淨銷售稅額（poslinev line_total_net＋line_tax）。
成交筆數：匯出沒有原單號，txn 的 原單日期 留空（已知 −0.13%，使用者 11/17 接受）。
只回傳 DataFrame 給引擎算門市彙總；員工名稱不讀（只留員工代碼，SA999 禮券回補判定要用）。
"""
import datetime as dt
import glob
import os
from pathlib import Path

import openpyxl
import pandas as pd

CACHE = Path(os.environ.get('EXPORT_CACHE') or Path.home() / '工具中心' / '匯入' / 'EPB匯出') / '_cache'
NEED = ['單據日期', '單據代碼', '交易類型', '門市代碼', '員工代碼', '存貨代碼', '名稱', '數量', '單位成本',
        '淨銷售金額(未稅)', '淨銷售稅額', '類別1代碼', '類別2代碼', '類別3代碼', '類別4代碼', '類別5代碼',
        '類別6代碼', '品牌代碼']


def _read(path):
    """一個匯出檔 → 原始欄位 DataFrame（字串）；快取成 csv，xlsx 沒變就直接讀快取。"""
    path = Path(path)
    CACHE.mkdir(parents=True, exist_ok=True)
    st = path.stat()
    cache = CACHE / f'{path.stem}__{st.st_size}_{int(st.st_mtime)}.csv'     # 換了同名新檔（大小／時間不同）就重建
    if cache.exists():
        return pd.read_csv(cache, dtype=str, keep_default_na=False)
    ws = openpyxl.load_workbook(path, read_only=True).worksheets[0]
    ws.reset_dimensions()
    rows, idx = [], None
    for r in ws.iter_rows(values_only=True):
        if idx is None:
            if sum(x not in (None, '') for x in r) >= 20:
                hdr = [str(x).strip() if x is not None else '' for x in r]
                miss = [c for c in NEED if c not in hdr]
                if miss:
                    raise ValueError(f'{path.name} 缺欄位 {miss}')
                idx = [hdr.index(c) for c in NEED]
            continue
        if r[idx[0]] in (None, ''):
            continue
        rows.append(['' if r[i] is None else str(r[i]) for i in idx])
    df = pd.DataFrame(rows, columns=NEED)
    tmp = cache.with_suffix('.tmp')
    df.to_csv(tmp, index=False)
    os.replace(tmp, cache)                       # 寫完才換上，中斷不會留下半份快取
    for old in CACHE.glob(f'{glob.escape(path.stem)}__*.csv'):
        if old != cache:
            old.unlink()
    return df


def meta(path):
    """匯出條件（中繼資料列）→ {'單據日期': 'YYYY-MM-DD, YYYY-MM-DD', '地點代碼': '024, 046, …', '導出時間': …}。"""
    ws = openpyxl.load_workbook(path, read_only=True).worksheets[0]
    ws.reset_dimensions()
    out = {}
    for r in ws.iter_rows(max_row=15, values_only=True):
        v = [x for x in r if x not in (None, '')]
        if v and len(v) < 6 and v[0] in ('單據日期', '地點代碼', '導出時間'):      # 表頭列第一欄也叫單據日期
            out[v[0]] = str(v[-1]).strip()
        elif 3 <= len(v) < 6:                     # 其他篩選條件列（例：類別1代碼 IN …）→ 只匯了部分品項
            out.setdefault('其他條件', []).append(str(v[0]).strip())
        elif len(v) >= 20:
            out['缺欄位'] = [c for c in NEED if c not in [str(x).strip() for x in v]]
            break
    return out


def check_file(path):
    """匯出檔能不能用：回 None＝可以，否則回原因（傳到一半、不是匯出檔、缺欄位）。"""
    try:
        m = meta(path)
    except Exception as e:
        return f'讀不到（{type(e).__name__}，可能還沒傳完）'
    if '單據日期' not in m or '地點代碼' not in m or '缺欄位' not in m:
        return '不是「銷售分析作業」匯出檔'
    if m['缺欄位']:
        return f'缺欄位 {"、".join(m["缺欄位"])}'
    if m.get('其他條件'):                         # 只匯部分品項的檔若當成完整資料，會蓋掉同日同店的完整檔
        return f'有其他篩選條件 {"、".join(m["其他條件"])}，不是完整資料'
    return None


def date_range(path):
    return meta(path).get('單據日期', '')


def _scope(path):
    """(起日, 迄日, 門市集合, 導出時間)；同一天同一店有好幾份檔時，用導出時間最新的那份。"""
    m = meta(path)
    a, b = [pd.Timestamp(x.strip()) for x in m['單據日期'].split(',')]
    return a, b, {s.strip().zfill(3) for s in m['地點代碼'].split(',')}, m.get('導出時間', '')


def _paths(paths):
    return sorted(glob.glob(str(paths))) if isinstance(paths, (str, Path)) else list(paths)


def _covered_days(paths):
    """{門市代碼: 有匯出檔涵蓋的日期集合}。匯出當天的資料還沒過完帳，不算涵蓋（之後補匯會用新檔蓋掉）。"""
    out = {}
    for a, b, st, t in (_scope(p) for p in _paths(paths)):
        if t:
            b = min(b, pd.Timestamp(t[:10]) - pd.Timedelta(days=1))
        days = {x.date() for x in pd.date_range(a, b)} if b >= a else set()
        for s in st:
            out.setdefault(s, set()).update(days)
    return out


def coverage(paths):
    """{門市代碼: 從最早一天起「連續」涵蓋到的最後一天(date)}；中間漏一天就停在漏的前一天。"""
    out = {}
    for s, days in _covered_days(paths).items():
        d = min(days) if days else None
        while d and d + dt.timedelta(days=1) in days:
            d += dt.timedelta(days=1)
        if d:
            out[s] = d
    return out


def missing_days(paths, start, end, store_codes):
    """[start, end] 期間每店沒有匯出檔涵蓋的日期 → {門市代碼: [date…]}（沒缺回空 dict）。"""
    cov = _covered_days(paths)
    want = [x.date() for x in pd.date_range(pd.Timestamp(str(start)[:10]), pd.Timestamp(str(end)[:10]))]
    out = {}
    for s in store_codes:
        lack = [d for d in want if d not in cov.get(s, set())]
        if lack:
            out[s] = lack
    return out


def select_raw(paths, start, end, store_codes=None):
    """多個匯出檔合併 → 原始欄位（字串）＋ 單據日期，限期間／門市。load() 與 epb_offline 的 SQLite 共用。"""
    paths = _paths(paths)
    # 不能 drop_duplicates：匯出沒有行號，同單同品項分開刷的兩列內容完全一樣（11/17 對帳少算 iPhone、西門 iPad 96→58）
    # 檔案重疊（每日單日檔＋補匯的整月檔）改用「同一天同一店只取導出時間最新的檔」避免重複計算
    scopes = [_scope(p) for p in paths]
    parts = []
    for i, p in enumerate(paths):
        df = _read(p)
        df['門市代碼'] = df['門市代碼'].str.strip().str.zfill(3)
        day = pd.to_datetime(df['單據日期'].str[:10], errors='coerce')
        keep = pd.Series(True, index=df.index)
        for j, (a, b, st, t) in enumerate(scopes):
            if j != i and (t > scopes[i][3] or (t == scopes[i][3] and j < i)):
                keep &= ~((day >= a) & (day <= b) & df['門市代碼'].isin(st))
        parts.append(df[keep])
    raw = pd.concat(parts, ignore_index=True)
    d = pd.to_datetime(raw['單據日期'].str[:19], errors='coerce')
    m = (d >= pd.Timestamp(start)) & (d <= pd.Timestamp(f'{end} 23:59:59'))
    if store_codes:
        m &= raw['門市代碼'].isin(store_codes)
    return raw[m], d[m]


def load(paths, start, end, store_codes=None):
    """多個匯出檔合併 → 引擎格式 DataFrame（同 load_from_epb 回傳），df.attrs['txn'] 也備好。"""
    raw, d = select_raw(paths, start, end, store_codes)
    num = lambda c: pd.to_numeric(raw[c], errors='coerce').fillna(0)
    numc = lambda c: pd.to_numeric(raw[c].replace('', None), errors='coerce')
    net, tax = num('淨銷售金額(未稅)'), num('淨銷售稅額')
    df = pd.DataFrame({
        '交易類型': raw['交易類型'].str.strip(),
        '單據日期': d,
        '單據代碼': raw['單據代碼'].str.strip(),
        '員工代碼': raw['員工代碼'].str.strip(),
        '員工名稱': raw['員工代碼'].str.strip(),
        '存貨代碼': raw['存貨代碼'].str.strip(),
        '名稱': raw['名稱'].str.strip(),
        '數量': num('數量'),
        '銷售金額(含稅)': net + tax,
        '銷退金額': 0.0,
        '淨銷售金額(未稅)': net,
        '銷退金額(未稅)': 0.0,
        '單位成本': num('單位成本'),
        '折扣': 0.0,
        '類別1代碼': numc('類別1代碼'),
        '類別2代碼': raw['類別2代碼'].str.strip(),
        '類別3代碼': numc('類別3代碼'),
        '類別4代碼': numc('類別4代碼'),
        '類別5代碼': numc('類別5代碼'),
        '類別6代碼': numc('類別6代碼'),
        '品牌代碼': numc('品牌代碼'),
        '地點代碼': raw['門市代碼'],
    }).reset_index(drop=True)
    df['NET'] = df['銷售金額(含稅)'] + df['銷退金額']
    tt = {'銷售': 'A', '尾款': 'H', '銷退': 'E'}
    t = df[df['交易類型'].isin(tt)].assign(日期=df['單據日期'].dt.strftime('%Y-%m-%d'))
    txn = (t.groupby(['地點代碼', '日期', '交易類型'])['單據代碼'].nunique().rename('張數').reset_index())
    txn['交易類型'] = txn['交易類型'].map(tt)
    txn['原單日期'] = ''
    df.attrs['txn'] = txn[['地點代碼', '日期', '交易類型', '原單日期', '張數']]
    return df
