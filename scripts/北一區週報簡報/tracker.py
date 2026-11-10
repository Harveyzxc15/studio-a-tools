# -*- coding: utf-8 -*-
"""週報評語追蹤帳本＋選題矩陣（寫評語前用）。

  python3 tracker.py 27Q1W06 [--region 北二區] [--dry]

1. metrics.json：從本週週報 Excel（BY店 本週比較／本週其他細項／重點項目進度）與評語素材（3PP 零搭售、MySetup）
   整理出每店的指標（本週、上週），存在 <輸出>/<週次>/metrics.json。
2. 追蹤帳本（<輸出>/追蹤帳本.json）：上週定稿每一格提到的主題＋要求動作記成一筆；本週用 metrics 判定
   改善（結案，只在別格帶一句）／持平／變差（續追，要換做法）。連續 3 週沒改善就提醒升級處理。
3. 選題矩陣：每店×品類的零搭售、AC+ 本週值與變化，標出問題格；3 店以上同一個問題＝區級議題（講稿講一次，評語最多兩店）；
   附各店近 4 週定稿已寫過的題目，避免重複。
帳本與矩陣寫進評語素材.md 的「## 追蹤帳本」「## 選題矩陣」兩節（重跑會覆蓋這兩節）。
主題判斷用 ../核對/topics.py（check.py 的跨格檢查也用同一份）。
"""
import argparse
import json
import re
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / '核對'))
import topics as T  # noqa: E402
import build_deck as B  # noqa: E402

PAIR_BLOCKS = {
    'iPhone 配件件數': {'保貼數量': 'film:iPhone', '保護殼數量': 'case:iPhone', '鏡頭貼數量': 'lens:iPhone'},
    'iPad 配件': {'原廠筆數量': 'pencil:iPad', '保護殼數量': 'case:iPad', '保貼數量': 'film:iPad', '鍵盤數量': 'kb:iPad'},
    'Watch 配件': {'錶帶數量': 'band:Watch', '保貼數量': 'film:Watch'},
    'CPU 配件': {'保貼數量': 'film:CPU', '電腦包數量': 'bag:CPU'},
}
AC_DEV = {'CPU': 'CPU', 'Watch': 'Watch', 'iPad': 'iPad', 'iPhone': 'iPhone', 'Airpods': 'AirPods', 'AirPods': 'AirPods'}
RATE_KEYS = ('zero', 'ac', 'coupon', 'mysetup')


def unit_of(key):
    k = key.split(':')[0]
    return 'rate' if k in RATE_KEYS else 'amt' if k in ('tpp', 'speaker') else 'cnt'


def fmt(key, v):
    if v is None:
        return '-'
    u = unit_of(key)
    if u == 'rate':
        return f'{v * 100:.1f}%'
    if u == 'amt':
        return f'{v / 10000:.1f} 萬' if abs(v) >= 10000 else f'{v:,.0f} 元'
    return f'{v:g}'


# ───────────── 指標 ─────────────
def _store(v):
    return B.store_key(v) if isinstance(v, str) else None


def _rows_below(ws, r, c):
    """(列, 店) ：從 r 往下，c 欄是店名的列，遇到 Total 停。"""
    out = []
    for rr in range(r, min(r + 15, ws.max_row + 1)):
        v = ws.cell(rr, c).value
        if v == 'Total':
            break
        s = _store(v)
        if s:
            out.append((rr, s))
    return out


def _find(ws, text, rows=200, cols=60, exact=True):
    for r in range(1, min(rows, ws.max_row) + 1):
        for c in range(1, min(cols, ws.max_column) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and (v.strip() == text if exact else text in v):
                yield r, c


def collect(work, need_prev=True):
    """{店: {鍵: {'cur': 值, 'prev': 值}}}"""
    mat = (work / '評語素材.md').read_text(encoding='utf-8')
    m = re.search(r'本週 (\d\d/\d\d)~(\d\d/\d\d)，上週 (\d\d/\d\d)~(\d\d/\d\d)', mat)
    cur_lab, prev_lab = f'{m[1]}~{m[2]}', f'{m[3]}~{m[4]}'
    xs = sorted(work.glob('tmp*/*.xlsx'), key=lambda f: f.stat().st_mtime) or \
        sorted(work.glob('*週報_*.xlsx'), key=lambda f: f.stat().st_mtime)
    if not xs:
        raise FileNotFoundError(f'{work} 沒有週報 Excel（重算暫存或 *週報_*.xlsx）')
    wb = openpyxl.load_workbook(xs[-1], data_only=True)
    M = {s: {} for s in B.NAMES}

    def put(s, key, cur=None, prev=None):
        if s in M:
            d = M[s].setdefault(key, {'cur': None, 'prev': None})
            if cur is not None:
                d['cur'] = cur
            if prev is not None:
                d['prev'] = prev

    ws = wb['BY店 本週比較']
    for r, c in _find(ws, '3PP(未加SA Care)', rows=5):
        for rr, s in _rows_below(ws, r + 2, 1):
            put(s, 'tpp', ws.cell(rr, c + 1).value, ws.cell(rr, c).value)
    for r, c in _find(ws, 'CPU', rows=40):
        if ws.cell(r, 1).value is None and ws.cell(r + 1, c).value and '~' in str(ws.cell(r + 1, c).value):
            for dev in ['CPU', 'iPad', 'iPhone', 'Watch', 'AirPods']:
                cc = next((x for x in range(c, c + 20) if ws.cell(r, x).value == dev), None)
                if cc:
                    for rr, s in _rows_below(ws, r + 2, 1):
                        put(s, f'units:{dev}', ws.cell(rr, cc + 1).value, ws.cell(rr, cc).value)
            break
    for r in range(1, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        if isinstance(lab, str) and lab.strip() in (cur_lab, prev_lab) and ws.cell(r, 2).value == 'CPU':
            which = 'cur' if lab.strip() == cur_lab else 'prev'
            for c in range(2, ws.max_column + 1):
                dev = AC_DEV.get(ws.cell(r, c).value)
                if dev and ws.cell(r, c + 2).value == 'ACPP+搭售率':
                    for rr, s in _rows_below(ws, r + 1, 1):
                        put(s, f'ac:{dev}', **{which: ws.cell(rr, c + 2).value})

    ws = wb['BY店 本週其他細項']
    for r, c in _find(ws, '活動使用'):
        cols = {ws.cell(r, x).value: x for x in range(c, c + 6)}
        for rr, s in _rows_below(ws, r + 1, c):
            g, x = ws.cell(rr, cols['抵用券贈出']).value, ws.cell(rr, cols['抵用券兌換']).value
            put(s, 'coupon', ((x or 0) / g) if g else None)
    for r, c in _find(ws, '喇叭銷售金額'):
        pc = next((x for x in range(c, c + 6) if ws.cell(r + 1, x).value == '上週'), None)
        cc = next((x for x in range(c, c + 6) if ws.cell(r + 1, x).value == '本週'), None)
        if not (pc and cc):
            continue
        for rr, s in _rows_below(ws, r + 2, c):
            put(s, 'speaker', ws.cell(rr, cc).value, ws.cell(rr, pc).value)
    for title, fields in PAIR_BLOCKS.items():
        for r, c in _find(ws, title):
            hdr = [(x, ws.cell(r + 1, x).value) for x in range(c, c + 40)]
            for f, key in fields.items():
                xs_ = [x for x, v in hdr if v == f]
                if len(xs_) >= 2:
                    for rr, s in _rows_below(ws, r + 2, c):
                        put(s, key, ws.cell(rr, xs_[1]).value, ws.cell(rr, xs_[0]).value)
    ws = wb['重點項目進度']
    for r, c in _find(ws, 'ARpedia'):
        if ws.cell(r + 1, c).value == '本週銷售':
            for x in range(c + 1, c + 12):
                s = _store(ws.cell(r, x).value)
                if s:
                    put(s, 'arpedia', ws.cell(r + 1, x).value)
            break

    sec = mat.split('## 3PP 搭售／零搭售', 1)[1].split('\n## ', 1)[0]
    for ln in sec.splitlines():
        p = [x.strip() for x in ln.split('|')]
        if len(p) == 4 and p[0] in M and '/' in p[2]:
            for which, val in (('cur', p[2]), ('prev', p[3])):
                a, _, z = (float(x) for x in val.split('/'))
                put(p[0], f'zero:{p[1]}', **{which: (z / (a + z)) if a + z else None})
                M[p[0]][f'zero:{p[1]}'][which + '_n'] = a + z
    ms = re.search(r'## MySetup 本週提交數 ／ 主機總數\n+(.+)', mat)
    if ms:
        for part in ms[1].split('|'):
            mm = re.match(r'\s*(\S+) (\d+)/(\d+)', part)
            if mm and mm[1] in M and int(mm[3]):
                put(mm[1], 'mysetup', int(mm[2]) / int(mm[3]))

    if need_prev:                                   # 只有本週值的指標，上週值從上週資料夾補
        prev_dir = prev_week_dir(work)
        if prev_dir and (prev_dir / '評語素材.md').exists():
            try:
                PM = collect(prev_dir, need_prev=False)
                for s in M:
                    for key in ('coupon', 'arpedia', 'mysetup'):
                        if key in PM.get(s, {}):
                            put(s, key, prev=PM[s][key]['cur'])
            except Exception as e:                  # 上週資料夾不完整就算了，不影響本週
                print(f'⚠️ 上週指標讀不到（{str(e)[:60]}）')
    return M


def prev_week_dir(work):
    """同一層、週次比本週小的最新資料夾（例：27Q1W06 → 27Q1W05）；週次資料夾名稱＝P.full。"""
    lab = work.name
    sib = sorted(d for d in work.parent.iterdir()
                 if d.is_dir() and re.fullmatch(r'\d\dQ\dW\d\d', d.name) and d.name < lab)
    return sib[-1] if sib else None


# ───────────── 帳本 ─────────────
def verdict(key, ref, now):
    if ref is None or now is None:
        return '？'
    k = key.split(':')[0]
    d = now - ref
    if unit_of(key) == 'rate':
        tol = 0.02
    elif unit_of(key) == 'amt':
        tol = max(abs(ref) * 0.05, 1000)
    else:
        tol = max(abs(ref) * 0.05, 1)
    if abs(d) < tol:
        return '持平'
    better = d < 0 if k in T.LOWER_BETTER else d > 0
    return '改善' if better else '變差'


def finals_history(n=4):
    """近 n 週定稿（評語範例/）→ [(週次, {店: [A, B]})]，新的在前。"""
    out = []
    for p in sorted(B.R['examples'].glob(f"{B.R['deck_prefix']} *.txt"), reverse=True)[:n]:
        lab = re.search(r'(\d\dQ\dW\d\d)', p.name)[1]
        d, cur = {}, None
        for ln in p.read_text(encoding='utf-8').splitlines():
            m = re.match(r'【(.+)】', ln)
            if m:
                cur = B.store_key(m[1]) or m[1]
                d[cur] = []
            elif cur and re.match(r'[AB][:：]', ln):
                d[cur].append(ln[2:].strip())
        out.append((lab, d))
    return out


def update_ledger(ledger, label, prev_label, finals, M):
    items = ledger.setdefault('items', [])
    # 還在追的＋本週（重跑時）才結案的都算「已有」，不重開
    find = {(i['store'], i['key']): i for i in items if i['status'] == 'open' or i.get('closed') == label}
    done = ledger.setdefault('finals_done', {})        # 定稿週次 → 處理它的那一週
    if prev_label and done.get(prev_label, label) != label:
        finals = {}                                     # 這份定稿之前就處理過了（例如這週沒收到新定稿），不再開帳
    elif prev_label and finals:
        done[prev_label] = label
    # 1) 上週定稿 → 開新項目或把動作加到既有項目
    for s, cells in finals.items():
        s = B.store_key(s) or s
        for cell in cells:
            act = T.split_action(cell)[1]
            for key in T.topics_of(cell):
                if key.startswith('units:'):
                    continue
                it = find.get((s, key))
                if not it:
                    base = M.get(s, {}).get(key, {}).get('prev')
                    it = {'store': s, 'key': key, 'opened': prev_label, 'baseline': base,
                          'polarity': T.polarity(cell, key), 'actions': [], 'history': [], 'status': 'open'}
                    items.append(it)
                    find[(s, key)] = it
                if act and not any(a['label'] == prev_label and a['text'] == act for a in it['actions']):
                    it['actions'].append({'label': prev_label, 'text': act})
    # 2) 本週結果
    for it in items:
        if it['status'] != 'open' and it.get('closed') != label:
            continue
        it['history'] = [h for h in it['history'] if h['label'] != label]       # 重跑同一週不重複記
        if it['history']:
            ref = it['history'][-1]['value']
        else:
            ref = M.get(it['store'], {}).get(it['key'], {}).get('prev')
            if ref is None:
                ref = it['baseline']
        now = M.get(it['store'], {}).get(it['key'], {}).get('cur')
        v = verdict(it['key'], ref, now)
        it['history'].append({'label': label, 'ref': ref, 'value': now, 'verdict': v})
        # 上週是肯定的題目（+）：沒有回落就結案；批評的題目：改善才結案
        if v == '改善' or (it.get('polarity') == '+' and v in ('持平', '？')):
            it['status'], it['closed'] = 'closed', label
        else:
            it['status'] = 'open'
            it.pop('closed', None)
    ledger['label'] = label
    return ledger


def ledger_md(ledger, label):
    rows = [i for i in ledger['items'] if i['history'] and i['history'][-1]['label'] == label]
    if not rows:
        return '## 追蹤帳本\n\n（上週定稿沒有可追蹤的主題）\n'
    order = {'變差': 0, '持平': 1, '？': 2, '改善': 3}
    rows.sort(key=lambda i: (order[i['history'][-1]['verdict']], B.NAMES.index(i['store']) if i['store'] in B.NAMES else 99))
    md = ['## 追蹤帳本（上週定稿提到的主題 → 本週結果）', '',
          '判定規則：改善＝結案，本週只在別格帶一句，不要再整格追；持平／變差＝續追，**要換一個沒用過的做法**；'
          '連續 3 週沒改善＝升級處理（店長約談、區級追蹤）。「當時的寫法」是上週定稿的方向（+ 肯定、- 批評）。', '',
          '| 店 | 主題 | 開立 | 當時的寫法 | 前值 | 本週 | 判定 | 已用過的要求（不要再用同一句） |', '|---|---|---|---|---|---|---|---|']
    for i in rows:
        h = i['history'][-1]
        streak = 0
        for x in reversed(i['history']):
            if x['verdict'] in ('持平', '變差'):
                streak += 1
            else:
                break
        tag = {'改善': '✅ 改善→結案', '持平': '➖ 持平→續追', '變差': '🔴 變差→續追', '？': '？ 沒有本週數字'}[h['verdict']]
        if i.get('polarity') == '+':
            tag = {'改善': '✅ 延續→結案', '持平': '✅ 維持→結案', '變差': '🔴 回落→可續追', '？': '？ 沒有本週數字'}[h['verdict']]
        if streak >= 3:
            tag += f'（連續 {streak} 週，升級處理）'
        acts = '；'.join(f'{a["label"][-3:]}：{a["text"][:40]}' for a in i['actions'][-3:])
        pol = {'+': '+', '-': '-', '=': '·'}[i.get('polarity', '=')]
        md.append(f'| {i["store"]} | {T.label(i["key"])} | {i["opened"][-3:]} | {pol} | {fmt(i["key"], h["ref"])} | '
                  f'{fmt(i["key"], h["value"])} | {tag} | {acts} |')
    return '\n'.join(md) + '\n'


# ───────────── 選題矩陣 ─────────────
def flags(M, s):
    out = {}
    for dev in T.DEVICES:
        z = M[s].get(f'zero:{dev}', {})
        if (z.get('cur_n') or 0) >= 5 and z.get('cur') is not None:
            dz = z['cur'] - z['prev'] if z.get('prev') is not None else 0
            if z['cur'] >= 0.5 or (dz >= 0.15 and z['cur'] >= 0.35):
                out[f'zero:{dev}'] = f'{dev} 零搭售 {z["cur"]*100:.0f}%' + (f'（{dz*100:+.0f}）' if dz else '')
        a, u = M[s].get(f'ac:{dev}', {}), M[s].get(f'units:{dev}', {}).get('cur') or 0
        if u >= 5 and a.get('cur') is not None:
            da = a['cur'] - a['prev'] if a.get('prev') is not None else 0
            if a['cur'] <= 0.2 or (da <= -0.15 and a['cur'] <= 0.35):
                out[f'ac:{dev}'] = f'{dev} AC+ {a["cur"]*100:.0f}%' + (f'（{da*100:+.0f}）' if da else '')
    for key, lim, name in (('speaker', -0.5, '喇叭'), ('tpp', -0.2, '3PP')):
        d = M[s].get(key, {})
        if d.get('cur') is not None and d.get('prev') and d['prev'] >= 5000:
            ch = d['cur'] / d['prev'] - 1
            if ch <= lim:
                out[key] = f'{name} {ch*100:+.0f}%'
    return out


DISTORT = ('大量採購', '送禮', '外國客', '教育價空機')     # 會讓零搭售／AC+ 偏高偏低的店長背景


def mgr_note(mgr, s, key):
    """問題格的店長背景：同品類（或沒寫品類）的大量採購／送禮／外國客／教育價空機。"""
    if ':' not in key:                                  # 喇叭、3PP 不標（背景多半跟品類有關）
        return ''
    dev = key.split(':')[1]
    kinds = []
    for f in mgr.get(s, {}).get('flags', []):
        if f['kind'] in DISTORT and dev in f.get('devices', []):
            if f['kind'] not in kinds:
                kinds.append(f['kind'])
    return f'（店長：{"、".join(kinds)}）' if kinds else ''


def adjust_bulk(M, bulk):
    """零搭售扣掉大單（bulk.json，bulk.cut_of 決定扣多少）→ (調整後的 M, {店: {鍵: 大單台數}})。上週值不動。"""
    if not bulk:
        return M, {}
    import bulk as BK
    M2, cut = json.loads(json.dumps(M)), {}
    for s, v in bulk['stores'].items():
        for dev, b in v['bulk'].items():
            z = M2.get(s, {}).get(f'zero:{dev}')
            if not z or z.get('cur') is None or not z.get('cur_n'):
                continue
            bn, bz = BK.cut_of(b, bulk['source'], z['cur_n'])
            n, zz = z['cur_n'] - bn, round(z['cur'] * z['cur_n']) - bz
            if not bn or n < 0 or zz < 0:         # 沒有要扣的，或對不上（簡報來源沒算到這些大單）就不扣
                continue
            z['cur'], z['cur_n'] = (zz / n if n else None), n
            cut.setdefault(s, {})[f'zero:{dev}'] = bn
    return M2, cut


def matrix_md(M, hist, mgr=None, bulk=None):
    mgr = mgr or {}
    M2, cut = adjust_bulk(M, bulk)
    F = {}
    for s in B.NAMES:
        raw, adj = flags(M, s), flags(M2, s)
        F[s] = {}
        for k, v in adj.items():
            F[s][k] = v + (f'（已排除大單 {cut[s][k]} 台）' if k in cut.get(s, {}) else '') + mgr_note(mgr, s, k)
        for k, v in raw.items():                # 原本是問題格、扣掉大單後就不是
            if k not in adj and k in cut.get(s, {}):
                z = M2[s][k]['cur']
                F[s][k + '#bulk'] = f'~~{v}~~ 排除大單 {cut[s][k]} 台後 {z*100:.0f}%，不算問題格' if z is not None \
                    else f'~~{v}~~ 全是大單，不算問題格'
    md = ['## 選題矩陣（本週問題格＋近 4 週定稿寫過的題目）', '',
          '問題格：零搭售 ≥50%（或增加 ≥15 個百分點且 ≥35%）、AC+ ≤20%（或下降 ≥15 個百分點且 ≤35%）（該品類本週 ≥5 台才算）、喇叭較前週 -50% 以下、3PP -20% 以下。'
          '零搭售已扣掉 EPB 辨識的大單（企業戶／通訊業、同客戶同品類 ≥5 台；簡報數字不變，評語可寫「含一筆大量採購」），較前週變化用扣掉後的本週值對上週。'
          '括號是較前週變化（百分點）；「店長：…」＝店長週報說有大量採購、送禮等背景，數字可能失真，寫之前先看「店長週報」一節。'
          '選題時先看追蹤帳本要續追的，再從這裡挑；同一個問題最多兩店。', '',
          '| 店 | 本週問題格 | 近 4 週定稿寫過 |', '|---|---|---|']
    for s in B.NAMES:
        seen = []
        for lab, d in hist:
            for cell in d.get(s, []):
                for k in T.topics_of(cell):
                    if not k.startswith('units:') and T.label(k) not in seen:
                        seen.append(T.label(k))
        md.append(f'| {s} | {"、".join(F[s].values()) or "—"} | {"、".join(seen) or "—"} |')
    cnt = {}
    for s, f in F.items():
        for k in f:
            if k.endswith('#bulk'):
                continue
            cnt.setdefault(k, []).append(s)
    # 店長說是大量採購造成的，不算進區級議題（例如遠百 AirPods 零搭售）
    cnt = {k: [s for s in ss if '大量採購' not in F[s][k]] for k, ss in cnt.items()}   # 大量採購造成的不算
    wide = [(k, ss) for k, ss in cnt.items() if len(ss) >= 3]
    md.append('')
    if wide:
        md.append('**區級議題**（3 店以上同一個問題 → 講稿講一次；評語只挑最嚴重的兩店寫，其他店換題）：')
        md += [f'- {T.label(k)}：{"、".join(ss)}' for k, ss in wide]
    else:
        md.append('區級議題：沒有（沒有 3 店以上同一個問題）')
    return '\n'.join(md) + '\n'


def write_sections(mat_path, sections):
    txt = mat_path.read_text(encoding='utf-8')
    for title, body in sections:
        txt = re.sub(rf'\n## {title}.*?(?=\n## |\Z)', '', txt, flags=re.S)
    head, sep, rest = txt.partition('\n## 上週快照')
    if not sep:
        head, sep, rest = txt.partition('\n## 3PP 搭售')
    new = head.rstrip() + '\n\n' + '\n'.join(b for _, b in sections) + sep + rest
    mat_path.write_text(new, encoding='utf-8')


def run(label, region, finals=None, dry=False, work=None):
    B.configure(region)
    work = Path(work) if work else B.OUT_BASE / label
    M = collect(work)
    (work / 'metrics.json').write_text(json.dumps(M, ensure_ascii=False, indent=1), encoding='utf-8')
    hist = finals_history()
    prev_label = next((lab for lab, _ in hist if lab < label), None)
    if finals is None:
        finals = dict(next((d for lab, d in hist if lab == prev_label), {}))
    lp = B.OUT_BASE / '追蹤帳本.json'
    ledger = json.loads(lp.read_text(encoding='utf-8')) if lp.exists() else {'region': region, 'items': []}
    ledger = update_ledger(ledger, label, prev_label or '', finals, M)
    if not dry:
        lp.write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding='utf-8')
    try:
        import manager_reports
        mgr = manager_reports.load(region, label)
    except Exception:
        mgr = {}
    import bulk as BK
    bk = BK.load(work)
    sections = [('追蹤帳本', ledger_md(ledger, label)), ('選題矩陣', matrix_md(M, hist, mgr, bk))]
    if bk:
        sections.append(('夥伴零搭售', BK.people_md(bk)))
    write_sections(work / '評語素材.md', sections)
    return ledger


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('label')
    ap.add_argument('--region', default='北一區')
    ap.add_argument('--dry', action='store_true', help='不寫入追蹤帳本.json（素材照寫）')
    a = ap.parse_args()
    run(a.label, a.region, dry=a.dry)
    B.configure(a.region)
    txt = (B.OUT_BASE / a.label / '評語素材.md').read_text(encoding='utf-8')
    print(re.search(r'## 追蹤帳本.*?(?=\n## 上週快照|\n## 3PP)', txt, re.S)[0])


if __name__ == '__main__':
    main()
