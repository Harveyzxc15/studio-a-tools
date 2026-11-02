"""簡報交件前的數字核對（build_deck 存檔後、fill_comments 匯出 PDF 前各跑一次）。

只看「做完的簡報」本身，不看程式以為寫了什麼：
  1. 回讀：本次寫入的每一格，Keynote 實際顯示的值要等於寫入值（抓 apply 沒寫進去、格式吃掉數字）
  2. 日期類：業績總表本月日數／應達成率、同頁日期小表、ARpedia 與重點項目的應達成率（本週／本月／年度）
  3. 公式錯誤：任何格子顯示 #REF!、#VALUE!、#NAME? 等（#DIV/0! 是目標為 0 的正常結果，不算）
  4. 週次：頁面文字裡的 Q#W## 都是本週

回傳 [(ok, 訊息)]；ok=False 的就是 ❌，交件前要改到沒有。
"""
import calendar
import re

PCT = re.compile(r'^-?[\d,]*\.?\d*%$')


def _parse(v):
    """顯示值或寫入值 → (數值, 容許誤差)；不是數字回 None。'-' 視為 0（業績總表 0 顯示成 -）。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v), 1e-9
    s = str(v).strip().replace(',', '')
    if s == '-':
        return 0.0, 1e-9
    m = re.fullmatch(r'(-?\d*\.?\d+)(%?)', s)
    if not m:
        return None
    dec = len(m[1].split('.')[1]) if '.' in m[1] else 0
    tol = 0.5 * 10 ** -dec
    if m[2]:
        return float(m[1]) / 100, tol / 100 + 1e-9
    return float(m[1]), tol + 1e-9


def same(shown, written):
    """Keynote 顯示的 shown 是否就是寫入的 written（依顯示的位數四捨五入比）。"""
    a, b = _parse(shown), _parse(written)
    if a is None or b is None:
        return str(shown).strip() == str(written).strip()
    return abs(a[0] - b[0]) <= max(a[1], b[1])


def _cell(t, r, c):
    return t['cells'][r][c][0].strip()


def _expect(res, where, shown, want, what):
    ok = same(shown, want)
    if not ok:
        res.append((False, f'{where}「{what}」顯示 {shown or "（空白）"}，應為 {want:.2%}' if isinstance(want, float)
                    else f'{where}「{what}」顯示 {shown or "（空白）"}，應為 {want}'))
    return ok


def verify(dump, P, ops=None, label=None):
    """dump：kn.js dump 的結果（list of slides）；P：build_deck.Period；ops：本次寫入的 ops（可省略）；
    label：週次文字（P.label）。"""
    res = []
    slides = {s['i']: s for s in dump}
    shown_slides = [s for s in dump if not s.get('skipped')]
    month = P.wk_end.day / P.days_in_month
    year = P.wk_end.timetuple().tm_yday / (366 if calendar.isleap(P.wk_end.year) else 365)
    n_dates = 0

    # 1. 回讀
    bad = 0
    for op in ops or []:
        if op.get('op') != 'cell' or op.get('v') in (None, ''):
            continue
        try:
            t = slides[op['s']]['tables'][op['t']]
            shown = _cell(t, op['r'], op['c'])
        except (KeyError, IndexError):
            res.append((False, f"第 {op['s'] + 1} 頁表格 {op['t']} 第 {op['r']},{op['c']} 格讀不到（版面可能被改）"))
            bad += 1
            continue
        if not same(shown, op['v']):
            bad += 1
            if bad <= 15:
                res.append((False, f"第 {op['s'] + 1} 頁「{t['name']}」第 {op['r'] + 1} 列第 {op['c'] + 1} 欄"
                                   f"寫入 {op['v']}，簡報顯示 {shown or '（空白）'}"))
    if ops is not None and not bad:
        n = sum(1 for op in ops if op.get('op') == 'cell' and op.get('v') not in (None, ''))
        res.append((True, f'回讀：本次寫入的 {n} 格都與簡報顯示一致'))
    elif bad > 15:
        res.append((False, f'回讀：另有 {bad - 15} 格不一致（只列前 15 格）'))

    for s in shown_slides:
        pg = f"第 {s['i'] + 1} 頁"
        for t in s['tables']:
            R, C = t['rows'], t['cols']
            # 2. 日期類
            if _cell(t, 0, 0).startswith('門市別') and C > 6:
                for r in range(R - 1):
                    if _cell(t, r, 6) == '本月日數':
                        _expect(res, pg, _cell(t, r + 1, 6), P.days_in_month, '本月日數')
                        n_dates += 1
                for r in range(R):
                    for c in range(C - 1):
                        if _cell(t, r, c) == '應達成率':
                            _expect(res, pg, _cell(t, r, c + 1), month, '業績總表 應達成率')
                            n_dates += 1
            if R == 3 and C >= 2 and _cell(t, 1, 0) == '日期' and _cell(t, 2, 0) == '應達成率':
                _expect(res, pg, _cell(t, 0, 1), P.days_in_month, '小表 天數')
                _expect(res, pg, _cell(t, 1, 1), P.wk_end.day, '小表 日期')
                _expect(res, pg, _cell(t, 2, 1), month, '小表 應達成率')
                n_dates += 3
            hdr = [_cell(t, 0, c) for c in range(C)]
            if '應達成率' in hdr and (_cell(t, 0, 0) in ('ARpedia', '重點項目')):
                k = hdr.index('應達成率')
                for r in range(1, R):
                    lab = _cell(t, r, 0)
                    want = {'本週銷售': 1.0, '年度銷售': year}.get(lab, month)
                    _expect(res, pg, _cell(t, r, k), want, f'{_cell(t, 0, 0)} {lab} 應達成率')
                    n_dates += 1
            # 3. 公式錯誤
            for r in range(R):
                for c in range(C):
                    v = _cell(t, r, c)
                    if v.startswith('#') and v != '#DIV/0!' and re.fullmatch(r'#[A-Z/0!?]+', v):
                        res.append((False, f"{pg}「{t['name']}」第 {r + 1} 列第 {c + 1} 欄顯示 {v}（公式錯誤）"))
        # 4. 週次
        if label:
            for x in s['texts']:
                for m in re.findall(r'Q\dW\d{1,2}', x['t']):
                    if m != label:
                        res.append((False, f'{pg}文字「{x["t"][:30]}」週次是 {m}，本週應為 {label}'))
    if n_dates and not any(not ok and '應為' in m and '週次' not in m for ok, m in res):
        res.append((True, f'日期類：{n_dates} 格（本月日數、應達成率等）都對（本月 {month:.2%}、年度 {year:.2%}）'))
    if not n_dates:
        res.append((False, '找不到業績總表／ARpedia 的應達成率格，版面可能改過，請人工確認'))
    return res


def report(res, title):
    lines = [f'# {title}', '']
    lines += [('✅ ' if ok else '❌ ') + m for ok, m in res if not ok] + [('✅ ' if ok else '❌ ') + m for ok, m in res if ok]
    n_bad = sum(1 for ok, _ in res if not ok)
    lines += ['', f'結果：{"沒有 ❌" if not n_bad else f"{n_bad} 項 ❌"}']
    return '\n'.join(lines) + '\n'
