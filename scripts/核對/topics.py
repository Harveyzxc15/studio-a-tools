# -*- coding: utf-8 -*-
"""週報評語的「主題」判斷與重複檢查（check.py 與週報工具的追蹤帳本 tracker.py 共用）。

主題鍵（metric key）：
  zero:<品類>  零搭售佔比（越低越好）      ac:<品類>  AC+ 搭售率
  units:<品類> 主機台數                    tpp  3PP 金額        speaker  喇叭金額
  arpedia  ARpedia 套數   coupon  抵用券兌換率   mysetup  MySetup 提交率
  film:<品類>  保貼  case:<品類> 保護殼  lens:iPhone 鏡頭貼  pencil:iPad 原廠筆
  kb:iPad 鍵盤  band:Watch 錶帶  bag:CPU 電腦包
品類：CPU iPad iPhone Watch AirPods
"""
import re
from difflib import SequenceMatcher

DEVICES = ['CPU', 'iPad', 'iPhone', 'Watch', 'AirPods']
_DEV = [('CPU', r'CPU|Mac(?!\s*mini\s*M4)|MacBook|電腦|MBP|MBA'), ('iPhone', r'iPhone|Duo'),
        ('iPad', r'iPad'), ('Watch', r'Watch|手錶'), ('AirPods', r'AirPods|耳機')]
# (正則, 鍵, 預設品類)；預設品類＝子句沒提到品類、前面也沒有時用
_METRIC = [
    (r'零搭售|未搭配件|沒有搭配件|未搭配配件|皆未搭', 'zero', None),
    (r'AC\+|ACPP|保固', 'ac', None),
    (r'喇叭|Speaker|音訊商品', 'speaker', ''),
    (r'3PP', 'tpp', ''),
    (r'ARpedia', 'arpedia', ''),
    (r'抵用券', 'coupon', ''),
    (r'MySetup', 'mysetup', ''),
    (r'錶帶', 'band', 'Watch'),
    (r'鏡頭貼', 'lens', 'iPhone'),
    (r'原廠筆|Pencil', 'pencil', 'iPad'),
    (r'鍵盤(?!膜)', 'kb', 'iPad'),
    (r'電腦包', 'bag', 'CPU'),
    (r'保貼|螢幕保護貼', 'film', 'iPhone'),
    (r'保護殼|殼貼|殼', 'case', 'iPhone'),
]
_UNITS = re.compile(r'銷售(?:由)?\s*\d+\s*[台隻支組]|由\s*\d+\s*[台隻支組](?:增加|減少|降)')
ACTION = re.compile(r'請(?:店長|門市|夥伴|導師|成交的夥伴)|店長每日|店長於')
LABEL = {'zero': '零搭售', 'ac': 'AC+', 'units': '台數', 'tpp': '3PP 金額', 'speaker': '喇叭', 'arpedia': 'ARpedia',
         'coupon': '抵用券兌換', 'mysetup': 'MySetup', 'film': '保貼', 'case': '保護殼', 'lens': '鏡頭貼',
         'pencil': '原廠筆', 'kb': '鍵盤', 'band': '錶帶', 'bag': '電腦包'}


def label(key):
    k, _, dev = key.partition(':')
    return f'{dev} {LABEL.get(k, k)}'.strip()


def split_action(text):
    """(敘述, 要求)：要求＝第一個「請店長／請門市／請夥伴…」之後。沒有就整段當敘述。"""
    m = ACTION.search(text)
    return (text[:m.start()], text[m.start():]) if m else (text, '')


def _dev_at(clause, last):
    """子句裡最先出現的品類（「iPhone 與 Mac」是 iPhone）；沒有就沿用前面的。"""
    hits = [(m.start(), d) for d, pat in _DEV for m in [re.search(pat, clause)] if m]
    return min(hits)[1] if hits else last


def clause_topics(text, with_units=True):
    """[(子句, [主題鍵…])]：品類沿用前面子句（「18 組未搭配件」接在「AirPods 銷售 34 組」後面就是 AirPods）。"""
    desc, _ = split_action(text)
    out, last = [], None
    for clause in re.split(r'[，。；]', desc):
        dev = _dev_at(clause, last)
        hits = []
        for pat, k, default in _METRIC:
            m = re.search(pat, clause)
            if not m:
                continue
            if default == '':
                key = k
            else:
                d = default if k in ('band', 'lens', 'pencil', 'kb', 'bag') else (dev or default)
                if not d:
                    continue
                key = f'{k}:{d}'
            hits.append((m.start(), key))
            if k in ('lens', 'film') and '殼貼' in clause:
                hits.append((m.start() + 1, f'case:{dev or "iPhone"}'))
        if with_units and dev and _UNITS.search(clause) and not hits:
            hits.append((0, f'units:{dev}'))
        keys = []
        for _, key in sorted(hits):
            if key not in keys:
                keys.append(key)
        out.append((clause, keys))
        if dev:
            last = dev
    return out


def topics_of(text, with_units=True):
    """依出現順序回傳主題鍵（不重複）。只看要求之前的敘述（要求裡的品項不算主題）。"""
    out = []
    for _, keys in clause_topics(text, with_units):
        out += [k for k in keys if k not in out]
    return out


_UP = re.compile(r'增加|提升|成長|回升|升到|升至|上升|拉高')
_DOWN = re.compile(r'減少|降|下滑|衰退|掉')
_BAD = re.compile(r'僅|未改善|仍有|仍未|掛零|尚未|沒有成交|較為可惜')
LOWER_BETTER = {'zero'}


def polarity(text, key):
    """這格對該主題是好轉（+）、變差（-）或持平／看不出（=）。看提到該主題的子句，第一個有升降字眼的為準。"""
    k = key.split(':')[0]
    ct = clause_topics(text, with_units=True)
    cls = []
    for i, (c, keys) in enumerate(ct):
        if key in keys:
            cls.append(c)
            if i + 1 < len(ct) and re.match(r'\s*(較|與|比|也較)', ct[i + 1][0]):   # 「3PP 完成 16.6 萬，較前週下滑 32.8%」
                cls[-1] += '，' + ct[i + 1][0]
    for clause in cls:
        up, down = bool(_UP.search(clause)), bool(_DOWN.search(clause))
        if up != down:
            return '+' if up != (k in LOWER_BETTER) else '-'
    if any(_BAD.search(c) for c in cls) or (cls and k in LOWER_BETTER):
        return '-'
    return '='


def main_topic(text):
    """這格的主題＝第一個非台數的主題鍵（都沒有才用台數）。"""
    t = topics_of(text)
    return next((k for k in t if not k.startswith('units:')), t[0] if t else None)


# ───────────── 重複檢查 ─────────────
_PRODUCT = re.compile(r'iPhone|iPad|AirPods|Watch|Speaker|ARpedia|MySetup|CPU|Mac|Duo|AC\+|3PP|DSS|SA Care')


def mask(s):
    """數字→#、產品名→品、去空白，讓「本週 iPhone 銷售 48 隻」和「本週 iPad 銷售 20 台」不會因數字不同被當成不一樣。"""
    s = _PRODUCT.sub('品', s)
    s = re.sub(r'[\d,.]+%?', '#', s)
    return re.sub(r'\s+', '', s)


def lcs(a, b):
    """最長共同連續片段（已遮罩）。回傳 (長度, 片段)。"""
    a, b = mask(a), mask(b)
    m = SequenceMatcher(None, a, b, autojunk=False).find_longest_match(0, len(a), 0, len(b))
    return m.size, a[m.a:m.a + m.size]


ACT_BAD, ACT_WARN, CELL_WARN = 10, 8, 18


def cross_checks(comments, last=None, source=None):
    """comments／last：{店: [A, B]}。回傳 [(狀態, 規則, 說明)]。"""
    res = []
    cells = [(f'{s}{"AB"[i]}', s, t) for s, v in comments.items() for i, t in enumerate(v) if t]
    olds = [(f'上週{s}{"AB"[i]}', s, t) for s, v in (last or {}).items() for i, t in enumerate(v) if t]

    def pairs(xs, ys, same_list):
        for i, x in enumerate(xs):
            for j, y in enumerate(ys):
                if same_list and j <= i:
                    continue
                yield x, y

    for xs, ys, same in [(cells, cells, True), (cells, olds, False)]:
        for (ca, _, ta), (cb, _, tb) in pairs(xs, ys, same):
            n, frag = lcs(split_action(ta)[1], split_action(tb)[1])
            if n >= ACT_BAD:
                res.append(('❌', '要求動作重複', f'{ca} 與 {cb} 的要求有 {n} 字相同：「{frag}」→ 換一個具體做法'))
            elif n >= ACT_WARN:
                res.append(('⚠️', '要求動作相近', f'{ca} 與 {cb}：「{frag}」'))
            else:
                n2, frag2 = lcs(ta, tb)
                if n2 >= CELL_WARN:
                    res.append(('⚠️', '句子重複', f'{ca} 與 {cb} 有 {n2} 字相同：「{frag2}」'))

    by_topic = {}
    for c, s, t in cells:
        k = main_topic(t)
        if k and polarity(t, k) == '-':
            by_topic.setdefault(k, set()).add(s)
    for k, ss in by_topic.items():
        if len(ss) > 2:
            res.append(('❌', '同一個問題超過兩店', f'{label(k)} 變差：{"、".join(sorted(ss))} → 第三店起換題，或改成區級議題放講稿'))
    for s, v in comments.items():
        if len(v) == 2 and all(v):
            ka, kb = main_topic(v[0]), main_topic(v[1])
            if ka and ka == kb:
                res.append(('❌', '同店兩格同主題', f'{s}：兩格都在講 {label(ka)}'))
            else:
                share = (set(topics_of(v[0])) & set(topics_of(v[1]))) - {'tpp'}
                for k in sorted(share):
                    nums = [set(re.findall(r'\d[\d,.]*\s*[台隻支組套片個條張筆萬元%]', ' '.join(
                        c for c, ks in clause_topics(t) if k in ks))) for t in v]
                    same = nums[0] & nums[1]
                    if same:
                        res.append(('❌', '同店兩格重提同一數字', f'{s}：兩格都寫 {label(k)}「{"、".join(sorted(same))}」'))
                    else:
                        res.append(('⚠️', '同店兩格重提', f'{s}：兩格都提到 {label(k)}'))
    opens = {}
    for c, s, t in cells:
        opens.setdefault(mask(t)[:6], []).append(c)
    for o, cs in opens.items():
        if len(cs) > 2:
            res.append(('⚠️', '開頭句型重複', f'{len(cs)} 格都用「{o}」開頭：{"、".join(cs)}'))
    if source is not None and '定稿信' not in source:
        res.append(('⚠️', '上週評語來源', f'不是信箱寄出的定稿（{source}）；比對與追蹤用的是自己上週的版本，請確認'))
    if not res:
        res.append(('✅', '重複與選題', '本週各格、與上週定稿都沒有重複'))
    return res


def last_finals_from_material(md):
    """評語素材.md 的「上週簡報的評語」區塊 → ({店: [A, B]}, 來源說明)。"""
    src = None
    m = re.search(r'上週評語來源：(.+)', md)
    if m:
        src = m[1].strip()
    sec = md.split('## 上週簡報的評語', 1)
    out = {}
    if len(sec) == 2:
        cur = None
        for ln in sec[1].split('\n## ', 1)[0].splitlines():
            if ln.startswith('### '):
                cur = ln[4:].strip()
                out[cur] = []
            elif cur and ln.startswith('- ') and '：' in ln:
                out[cur].append(ln.split('：', 1)[1].strip())
    return out, src
