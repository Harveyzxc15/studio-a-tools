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
SAME_STORE_BAD, SAME_STORE_WARN = 0.50, 0.42      # 同店本週 vs 上週（遮罩後整段相似度）
OPEN_SHARE = 0.5                                  # 同一種開頭佔全區格數超過這個比例就提醒
_OPEN = [('接續上週', r'上週(?:提醒|要求|追蹤)|連續[兩二三]週'),
         ('台數', r'銷售|由#[台隻支組套個](?:增加|減少|降)'),
         ('金額', r'完成(?:約)?#|營業額|業績|由#[萬元]'),
         ('問題', r'零搭售|未搭|沒有搭|下滑|未改善|偏低')]
OPEN_LABEL = {'接續上週': '「上週提醒／要求…」', '台數': '主機台數（本週 X 銷售…）', '金額': '金額（本週 X 完成…）',
              '問題': '直接點問題', '其他': '其他'}


def opening_type(text):
    """第一個子句屬於哪種開頭：接續上週／台數／金額／問題／其他。"""
    first = re.split(r'[，。；]', mask(text), 1)[0]
    return next((name for name, pat in _OPEN if re.search(pat, first)), '其他')


def ratio(a, b):
    """遮罩（數字→#、產品→品）後的整段相似度：骨架一樣、只換數字和品類也算像。"""
    return SequenceMatcher(None, mask(a), mask(b), autojunk=False).ratio()


def week_checks(comments, last=None, older=None):
    """同一家店跨週：跟上週太像、連續幾週同主題、開頭寫法太集中。older＝上上週 {店: [A, B]}。"""
    res = []
    for s, v in comments.items():
        prev = [t for t in (last or {}).get(s, []) if t]
        old = [t for t in (older or {}).get(s, []) if t]
        for i, t in enumerate(v):
            if not t:
                continue
            cell = f'{s}{"AB"[i]}'
            sim, j = max(((ratio(t, o), j) for j, o in enumerate(prev)), default=(0, 0))
            if sim >= SAME_STORE_BAD:
                res.append(('❌', '同店跟上週太像', f'{cell} 與上週{s}{"AB"[j]} 相似 {sim:.0%}（骨架一樣只換數字）→ 換切入角度與句型重寫'))
            elif sim >= SAME_STORE_WARN:
                res.append(('⚠️', '同店跟上週相近', f'{cell} 與上週{s}{"AB"[j]} 相似 {sim:.0%}'))
            k = main_topic(t)
            if k and any(main_topic(o) == k for o in prev):
                if any(main_topic(o) == k for o in old):
                    res.append(('⚠️', '同店連三週同主題', f'{cell}：{label(k)} 已連續第三週當主題 → 已改善就換題（併進另一格一句帶過）；仍落後要續追的話，角度、數字、要求動作全部換新'))
                else:
                    res.append(('⚠️', '同店連兩週同主題', f'{cell}：{label(k)} 上週也是主題 → 可以追蹤，但切入角度、對比的數字、要求動作都要換'))
        for i, t in enumerate(v):
            if t and i < len((last or {}).get(s, [])) and last[s][i] and opening_type(t) == opening_type(last[s][i]) != '其他':
                res.append(('⚠️', '同店同格開頭跟上週一樣', f'{s}{"AB"[i]}：連兩週都用{OPEN_LABEL[opening_type(t)]}開頭'))
    cells = [(f'{s}{"AB"[i]}', opening_type(t)) for s, v in comments.items() for i, t in enumerate(v) if t]
    for name in OPEN_LABEL:
        cs = [c for c, o in cells if o == name]
        if name != '其他' and len(cells) >= 6 and len(cs) > len(cells) * OPEN_SHARE:
            res.append(('⚠️', '開頭寫法太集中', f'{len(cs)}/{len(cells)} 格都用{OPEN_LABEL[name]}開頭：{"、".join(cs)} → 至少一半改用其他開頭（先講結論、先講對比、先講上週追蹤結果、先講問題）'))
    return res


def cross_checks(comments, last=None, source=None, older=None):
    """comments／last／older：{店: [A, B]}（本週／上週定稿／上上週）。回傳 [(狀態, 規則, 說明)]。"""
    res = week_checks(comments, last, older)
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


SPEECH_TYPES = ['數字總覽', '主軸先行', '門市故事', '追蹤回顧', '月進度']


def speech_type(md):
    """講稿.md 標的結構（標題或任一行寫「結構：主軸先行」）；沒標回 None。"""
    m = re.search(r'結構[：:]\s*(' + '|'.join(SPEECH_TYPES) + ')', md)
    return m[1] if m else None


def _speech_open(md):
    body = [l.strip() for l in md.splitlines() if l.strip() and not l.lstrip().startswith('#')]
    return re.split(r'[。；]', body[0], 1)[0] if body else ''


def speech_checks(md, last_md=None):
    """週報講稿跨週多元化（2026-11-13 加）：要標結構、不能連兩週同一種、開場句不能跟上週同骨架。"""
    res, kind = [], speech_type(md)
    if not kind:
        res.append(('⚠️', '講稿沒標結構', f'標題加「結構：{"／".join(SPEECH_TYPES)}」其中一種（見 講稿修改紀錄.md「講稿多元化」）'))
    if last_md:
        last_kind = speech_type(last_md) or '數字總覽'          # 舊講稿沒標的都是數字總覽型
        if (kind or '數字總覽') == last_kind:
            res.append(('❌', '講稿結構跟上週一樣', f'上週是「{last_kind}」，這週換另一種結構'))
        n, frag = lcs(_speech_open(md), _speech_open(last_md))
        if n >= 10:
            res.append(('⚠️', '講稿開場句跟上週同骨架', f'「{frag}」→ 換開場（先講主軸、先講一家店、先講上週追蹤結果）'))
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
