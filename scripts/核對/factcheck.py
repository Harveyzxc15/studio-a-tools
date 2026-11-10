# -*- coding: utf-8 -*-
"""交付前核對：把草稿裡的每個數字對回「事實檔」，再跑寫法規則，產出核對單。

事實檔（facts.json）＝這次報告能用的全部數字，由 daily_facts.py／weekly_facts.py／achv_facts.py 產生，
臨時查的數字用 `check.py add` 補進去（要寫來源）。草稿裡任何數字對不到事實檔就是 ❌，不能交。

事實格式：{"v": 數值, "u": "rate"|"amt"|"cnt"|null, "s": 門市或"", "l": 標籤, "g": 群組, "src": 來源}
  rate 存小數（27% → 0.27）。同一個 g 的事實會自動衍生差額、增減%、比值、百分點差，
  所以「由 7 台增加至 16 台、增加 9 台、成長 128.6%」只要 7 跟 16 在同一群組就對得到。
"""
import bisect
import json
import math
import re
from itertools import combinations

N1_STORES = ['士林', '微風', '美麗華', '阿波羅', '高島屋', '羅東']
N2_STORES = ['永和', '板橋誠品', '誠品', '西門', '花蓮', '板橋遠百', '遠百', '新莊宏匯', '宏匯', '裕隆城']
STORE_ALIAS = {'大葉高島屋': '高島屋', '誠品': '板橋誠品', '遠百': '板橋遠百', '宏匯': '新莊宏匯',
               '新西門': '西門', '新店裕隆城': '裕隆城'}
REGION_WORDS = ['北一區', '北ㄧ區', '北二區', '整區', '區內', '全區']
METRIC_WORDS = ['CPU', 'Mac mini', 'MacBook', 'iPad', 'iPhone', 'Duo', 'Watch', 'AirPods', 'ARpedia', 'MySetup',
                '3PP', '喇叭', '音頻', '保固', 'AC+', 'SA Care', 'ACPP', '營業額', '業績', '毛利', '來客', '成交',
                '舊換新', '配件', '保護殼', '保貼', '鏡頭貼', '抵用券', '禮券', '零搭售', '人流', '提袋']

UNITS = '%％萬億台隻支套組筆片顆位天元人次堂件個家間倍週成'
CNT_UNITS = set('台隻支套組筆片顆位天人次堂件個家間週')
NUM_RE = re.compile(r'(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d+))?(\s*)([' + UNITS + r'])?')
SKIP_AFTER = re.compile(r'^\s*(月|日|號|點|年|吋|寸|代|核|GB|TB|G\b|mm|小時|分鐘|:|/|－|-\d|Q\d|W\d|[A-Za-z])')
SKIP_BEFORE = re.compile(r'(第|Q|W|/|:|~|～|v|V|No\.?|#)\s*$')
MODEL_BEFORE = re.compile(r'[A-Za-z+]\s?$')


# ---------------------------------------------------------------- 事實檔
def load_facts(paths):
    facts = []
    for p in paths:
        with open(p, encoding='utf-8') as f:
            facts += json.load(f)['facts']
    for x in facts:
        if x.get('s') in REGION_WORDS:
            x['s'] = ''
    return facts


def save_facts(path, facts, meta=None):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({'meta': meta or {}, 'facts': facts}, f, ensure_ascii=False, indent=0)


def fact(v, u, s, l, g=None, src='', p=None):
    s = '' if (s or '') in REGION_WORDS else (s or '')     # 區合計門市欄＝空白（寫成「北二區」會被當成門市）
    return {'v': float(v), 'u': u, 's': s, 'l': l, 'g': g, 'src': src, 'p': p}


# 同一店同指標，哪些期間兩兩比較有意義（沒標期間的群組＝全部兩兩比）
PERIOD_PAIRS = {frozenset(['月累', '目標']), frozenset(['月累', '上月同期']), frozenset(['本週', '上週']),
                frozenset(['今日', '月累'])}


def derive(facts):
    """同群組兩兩衍生：差額、增減%、比值、百分點差。回傳 (值, 單位, 門市, 說明)。"""
    groups = {}
    for f in facts:
        if f.get('g'):
            groups.setdefault(f['g'], []).append(f)
    out = []
    # 任何「達成率」都自動算跟本月／年度應達成率的百分點差（落後／超前應達成 X%）
    paces = [f for f in facts if f['l'] in ('本月應達成率', '年度應達成率')]
    for f in facts:
        if f['u'] == 'rate' and '達成率' in f['l'] and '應達成率' not in f['l']:
            for p in paces:
                out.append((abs(f['v'] - p['v']), 'rate', f['s'], f"{f['l']} vs {p['l']}（落後／超前應達成）"))
    for members in groups.values():
        for a, b in combinations(members, 2):
            va, vb = a['v'], b['v']
            pair = frozenset([a.get('p'), b.get('p')])
            u = a['u'] if a['u'] == b['u'] else None
            desc = f"{a['l']} vs {b['l']}"
            if a['s'] != b['s']:
                # 跨店只允許「店 vs 區」的百分點差（例：士林 iPhone 保固比區內高 3%）
                if (a['s'] and b['s']) or u != 'rate':
                    continue
                out.append((abs(va - vb), 'rate', a['s'] or b['s'], f'百分點差 {desc}'))
                continue
            s = a['s']
            if a.get('p') and b.get('p') and pair not in PERIOD_PAIRS:
                continue
            if u == 'rate':
                out.append((abs(va - vb), 'rate', s, f'百分點差 {desc}'))
                continue
            out.append((abs(va - vb), u, s, f'差額 {desc}'))
            if pair == frozenset(['月累', '目標']):
                continue                      # 月累 vs 目標：達成率另有事實，不算增減%
            for x, y in ((va, vb), (vb, va)):
                if y:
                    out.append((abs(x / y - 1), 'rate', s, f'增減% {desc}'))
                    out.append((x / y, 'ratio', s, f'比值 {desc}'))
    return out


class Index:
    """把事實＋衍生值依數值排序，查詢 O(log n)。"""

    def __init__(self, facts):
        facts = [dict(f, s='') if f.get('s') in REGION_WORDS else f for f in facts]   # 區合計門市欄＝空白
        rows = [(f['v'], f['u'], f['s'], f['l'] + (f"［{f['src']}］" if f.get('src') else ''), True) for f in facts]
        rows += [(v, u, s, d, False) for v, u, s, d in derive(facts)]
        rows.sort(key=lambda r: r[0])
        self.rows = rows
        self.keys = [r[0] for r in rows]

    def between(self, lo, hi):
        i = bisect.bisect_left(self.keys, lo - 1e-9)
        j = bisect.bisect_right(self.keys, hi + 1e-9)
        return self.rows[i:j]


# ---------------------------------------------------------------- 抽數字
def extract_numbers(text):
    """抽出要核對的數字；日期、時間、型號、季週代號跳過。回傳 dict 清單（含位置與上下文）。"""
    out = []
    for m in NUM_RE.finditer(text):
        whole, dec, _sp, unit = m.group(1), m.group(2), m.group(3), m.group(4)
        before, after = text[max(0, m.start() - 12):m.start()], text[m.end():m.end() + 6]
        after_num = text[m.end(2) if dec else m.end(1):][:6]
        if SKIP_BEFORE.search(before) or (not unit and SKIP_AFTER.search(after_num)):
            continue
        if not unit and re.match(r'^[A-Za-z]', after_num):
            continue
        if unit is None and MODEL_BEFORE.search(before):
            continue                          # iPhone 18、M5、AirPods 4 這類型號
        if re.search(r'[A-Za-z]$', before) and not _sp and not unit:
            continue
        if unit in ('天', '週') and re.match(r'^(內|後|前|起)', after):
            pass
        out.append({'text': m.group(0).strip(), 'x': float(whole.replace(',', '') + ('.' + dec if dec else '')),
                    'dec': len(dec) if dec else 0, 'unit': unit, 'comma': ',' in whole,
                    'pos': m.start(), 'end': m.end()})
    return out


PRODUCTS = {'CPU': ['CPU', 'Mac'], '電腦': ['CPU', 'Mac', '電腦'], 'Mac mini': ['Mac mini'], 'MacBook': ['MacBook', 'CPU'], 'iPad': ['iPad'],
            'iPad mini': ['iPad mini'], 'AirPods 4': ['AirPods 4'], 'MacBook Pro': ['MacBook Pro', 'MBP'],
            'iPhone': ['iPhone'], 'Duo': ['Duo'], 'Watch': ['Watch'], 'AirPods': ['AirPods']}
ATTRS = {'保固': ['保固', '合計', 'AC+'], 'AC+': ['AC+', 'AC'], 'ACPP': ['AC+', 'AC', 'ACPP', '保固'], 'SA Care': ['SA'],
         '3PP': ['3PP'], '營業額': ['營業額', '業績'], '業績': ['營業額', '業績'], '毛利': ['毛利'],
         '來客': ['來客', '人流'], '人流': ['來客', '人流'], '成交筆數': ['成交'], '舊換新': ['舊換新'],
         '零搭售': ['零搭售'], 'ARpedia': ['ARpedia'], 'MySetup': ['MySetup'], '喇叭': ['喇叭', 'Speaker'],
         '音頻': ['喇叭', '音頻'], 'Speaker': ['喇叭', 'Speaker'], '抵用券': ['抵用券', '兌換', '贈出'], '禮券': ['禮券'], '保護殼': ['保護殼'], '保貼': ['保貼', '保護貼'],
         '鏡頭貼': ['鏡頭貼'], '配件': ['配件', '週邊'], 'Apple 配件': ['Apple配件', 'Apple 配件'], '提袋': ['提袋']}


def _last_of(words, seg):
    best, at = None, -1
    for w in words:
        i = seg.rfind(w)
        if i >= 0 and i + len(w) > at:
            best, at = w, i + len(w)
    return best


STANDALONE = {'ARpedia', 'MySetup', '喇叭', 'Speaker', '音頻', '抵用券', '禮券', '3PP', '營業額', '業績', 'Apple 配件', '毛利'}


def context(text, pos, stores):
    """數字前面最近的門市（同子句優先，再看同一句）、產品、指標詞（同一句內）。"""
    sent_start = max(text.rfind(c, 0, pos) for c in '。\n') + 1
    clause_start = max([text.rfind(c, 0, pos) for c in '，,；;。\n、']) + 1
    sent, clause = text[sent_start:pos], text[clause_start:pos]
    names = stores + list(STORE_ALIAS) + REGION_WORDS
    in_clause = _last_of(names, clause)
    store = in_clause or _last_of([n for n in names if n != '區內'], sent)
    owner = re.match(r'^(\S{2,4})：', text)                 # 週報評語：整段都是這家店
    if not store and owner:
        store = owner[1]                      # 段落分成好幾句（Yuni 寫法）：後面的句子沒講店名＝還是這家店
    store = STORE_ALIAS.get(store, store)
    if not store or store in REGION_WORDS:
        store = ''                            # 沒講店名＝講全區
    # 子句沒講店名時，前面講過的店不一定延續（Bob 整段逗號串接）→ 該店或全區都接受
    owned = owner is not None
    stores_ok = {store} if (in_clause or owned) else {store, ''}
    prod_clause = _last_of(list(PRODUCTS), clause)
    product, attr = prod_clause or _last_of(list(PRODUCTS), sent), _last_of(list(ATTRS), sent)
    if attr in STANDALONE and product and sent.rfind(attr) > sent.rfind(product):
        product = None                        # 「iPad 7 台，ARpedia 1 套」：ARpedia 不屬於 iPad
    elif attr in STANDALONE and product and sent.rfind(product) > sent.rfind(attr):
        attr = None                           # 「Speaker 3.8 萬，iPad 17 台，增加 4 台」：後面講的 iPad 才是主詞
    loose_prod = product is not None and prod_clause is None   # 產品是前面子句延續來的，不一定還適用
    loose_attr = attr is not None and _last_of(list(ATTRS), clause) is None
    return store, product, attr, (stores_ok, loose_prod, loose_attr, prod_clause is not None)


# ---------------------------------------------------------------- 比對
def candidates(tok, idx):
    x, d, u = tok['x'], tok['dec'], tok['unit']
    step = 10 ** -d
    hits = []

    def add(lo, hi, allowed):
        for v, fu, s, l, base in idx.between(lo, hi):
            if fu in allowed:
                hits.append((v, fu, s, l, base))

    if u in ('%', '％'):
        # 四捨五入或無條件捨去都算對；百分點差、增減% 也是 rate
        add((x - step / 2) / 100, (x + step) / 100, {'rate', None})
        if x > 1.5 and d == 0:
            add(x - 0.5, x + 0.5, {'pct'})    # 事實本身就是 % 數（素材寫 27%）
    elif u == '成':
        add((x - 0.5) / 10, (x + 0.5) / 10, {'rate', None})   # 「成長近 9 成」＝約 90%
    elif u == '倍':
        add(x - step / 2, x + step, {'ratio', None})
    elif u in ('萬', '億'):
        k = 1e4 if u == '萬' else 1e8
        add((x - step / 2) * k, (x + step) * k, {'amt', None})
    elif u == '元' or tok['comma'] or x >= 10000:
        add(x - 1, x + 1, {'amt', None, 'cnt'})
    elif u in CNT_UNITS:
        if d:
            add(x - step / 2, x + step / 2, {'cnt', None, 'avg'})
        else:
            add(x, x, {'cnt', None})
    else:
        add(x - step / 2, x + step / 2, {'cnt', 'amt', None, 'avg'})
        if d:
            add((x - step / 2) / 100, (x + step / 2) / 100, {'rate'})
    return hits


TARGET_AFTER = re.compile(r'^\s*(以上|以下|目標|為目標|的目標)')
TARGET_BEFORE = re.compile(r'(目標|穩定在|維持在|拉高到|拉高至|提升到|往|達到|超過|訂下的?)\s*$')
DIFF_WORDS = re.compile(r'增加|減少|多|少|差|成長|下滑|衰退|提升|下降|增|減|落後|超前|高出|高於|低於|相差|差距|較|升|降|超過|不到|僅差|剩')
UNIT_FAMILY = {**{c: '台' for c in '台隻支'}, **{c: '套' for c in '套組'}, **{c: '件' for c in '件片顆個'}}
ACC_LABEL = re.compile(r'配件|數量|件數|膜|保護殼|保貼|鏡頭貼|筆')
ACC_ATTR = {'配件', '保護殼', '保貼', '鏡頭貼', 'Apple 配件', '零搭售'}


def check_numbers(text, facts, stores):
    idx = Index(facts)
    results = []
    last_end = 0
    for tok in extract_numbers(text):
        hits = candidates(tok, idx)
        store, product, attr, (stores_ok, loose_prod, loose_attr, prod_here) = context(text, tok['pos'], stores)
        _ss = max(text.rfind(c, 0, tok['pos']) for c in '。\n') + 1
        prods_sent = [k for k in PRODUCTS if k in text[_ss:tok['pos']]]
        attr0 = attr
        if tok['unit'] in CNT_UNITS and attr in ('營業額', '業績', '毛利'):
            attr = None                       # 「銷售 277 隻」是台數，前面講的營業額不算
        if tok['unit'] in ('台', '隻', '支', '組') and attr in ('保固', '零搭售'):
            attr = None                       # 台／隻／支＝主機台數；保固用「套」
        metric = '／'.join(x for x in (product, attr) if x)
        status, why = '❌', '事實檔裡找不到'
        fam = UNIT_FAMILY.get(tok['unit'])
        if fam:   # 文字事實帶單位（〔套〕）：台、套、件不互通
            hits = [h for h in hits if not re.search(r'〔(.)〕', h[3]) or
                    UNIT_FAMILY.get(re.search(r'〔(.)〕', h[3]).group(1)) in (fam, None)]
        if tok['unit'] in ('台', '隻', '支') and attr0 not in ACC_ATTR:
            hits = [h for h in hits if not ACC_LABEL.search(h[3].split('［')[0])]   # 主機台數不對配件件數
        seg = text[max([text.rfind(c, 0, tok['pos']) for c in '，,；;。\n、'] + [last_end]) + 1:tok['pos']]
        last_end = tok['end'] - 1
        per = _last_of(['上週', '前週', '本週'], seg)   # 只看「上一個數字到這個數字」之間
        clause = text[max([text.rfind(c, 0, tok['pos']) for c in '，,；;。\n、']) + 1:tok['pos']]
        if not DIFF_WORDS.search(clause):
            # 前面沒有「增加／減少／成長／差距…」就不能用衍生的差額、增減%、百分點差來對
            hits = [h for h in hits if h[4] or h[3].startswith('比值')]
        if per and hits:
            want, avoid = ('上週', '本週') if per in ('上週', '前週') else ('本週', '上週')
            kept = [h for h in hits if not (avoid in h[3] and want not in h[3])]
            if not kept:
                status, why = '⚠️', f'期間不符：句子講{per}，但對到的是{avoid}的數字 → {hits[0][3]}'
            hits = kept
        if hits:
            def fit(h):
                v, fu, s, l, base = h
                ok_store = s in stores_ok
                generic = l.startswith(('本月應達成率', '年度應達成率', '本月已過天數'))   # 跟品項無關
                no_prod = not any(w.lower() in l.lower() for ws in PRODUCTS.values() for w in ws)
                prod_hit = bool(product) and any(w.lower() in l.lower() for w in PRODUCTS[product])
                # 指標詞是前面子句延續來的：同子句有講產品、且產品對上就好（例「AirPods 銷售 153 台」）
                ok_attr = generic or not attr or any(w in l for w in ATTRS[attr]) \
                    or (loose_attr and prod_here and prod_hit)
                # 產品是前面子句延續來的、標籤本身沒產品：指標詞要明確對上才算（例「AC+ 金額 99 萬」）
                # 產品是前面子句延續來的：同一句前面講過的其他產品也算（「iPhone 304 支，其中 Duo 68 支…搭售率 24%」）
                ok_prod = generic or not product or any(w.lower() in l.lower() for w in PRODUCTS[product]) \
                    or (loose_prod and any(w.lower() in l.lower() for k in prods_sent for w in PRODUCTS[k])) \
                    or (loose_prod and no_prod and attr is not None and ok_attr)
                return (ok_store and ok_prod and ok_attr, ok_store + ok_prod + ok_attr, base)

            # 同分時挑最像的那筆：子句講的店（沒講＝全區）、期間詞對得上、% 對「率」、數值最接近
            # 「20% 目標」「目標 7.8%」才是目標值；「低於目標」只是比較
            is_target = re.match(r'^\s*(的)?目標', text[tok['end']:]) or re.search(r'目標\s*$', text[:tok['pos']])
            want_p = '目標' if is_target else ('上月同期' if '上月' in clause else None)
            want_s = store if len(stores_ok) == 1 else ''
            t = {'%': tok['x'] / 100, '％': tok['x'] / 100, '成': tok['x'] / 10,
                 '萬': tok['x'] * 1e4, '億': tok['x'] * 1e8}.get(tok['unit'], tok['x'])

            def tiebreak(h):
                v, fu, s, l, base = h
                head = l.split('［')[0]
                p_ok = (want_p in head) if want_p else not re.search(r'上月同期|目標|佔區|上週', head)
                rate_ok = tok['unit'] not in ('%', '％') or '率' in head
                tt = tok['x'] if fu == 'pct' else t
                return (s == want_s, p_ok, rate_ok, -abs(v - tt) / max(abs(tt), 1e-9))
            hits.sort(key=lambda h: (fit(h), tiebreak(h)), reverse=True)
            best = hits[0]
            if fit(best)[0]:
                store = best[2]                   # 上下文欄顯示實際對到的門市（沒講店名＝全區）
            if fit(best)[0]:
                status, why = '✅', best[3]
            else:
                status = '⚠️'
                why = f'數值對得到但門市／指標不一致 → {best[3]}'
        if status != '✅' and (TARGET_AFTER.search(text[tok['end']:tok['end'] + 6])
                               or TARGET_BEFORE.search(text[max(0, tok['pos'] - 8):tok['pos']])):
            status, why = '🎯', '自訂目標值（沒有資料來源，確認合理：跟現況、上月、目標表比）'
        results.append({'tok': tok, 'status': status, 'why': why, 'store': store, 'metric': metric,
                        'n_hits': len(hits)})
    return results


# ---------------------------------------------------------------- 寫法規則
def rule_results(text, profile, speech=False):
    """profile: daily_n1 / daily_n2 / weekly_n1 / weekly_n2 / analysis。回傳 [(狀態, 規則, 說明)]。
    speech＝週報講稿（.md，唸的稿子）：評語的「每格」格式規則（北二每格字數／句號、北一句尾無句號）不適用。"""
    out = []
    paras = [p for p in text.split('\n') if p.strip()]

    def bad(words, label, level='❌'):
        for w in words:
            if w in text:
                out.append((level, label, f'出現「{w}」'))

    if profile in ('weekly_n1', 'analysis'):
        bad(['動能', '拉回', '節奏'], '禁用詞（store-report-style）')
    if profile == 'daily_n1':
        bad(['動能', '拉回', '節奏'], '禁用詞（store-report-style；Bob 日報偶爾用，確認是否必要）', '⚠️')
    if profile in ('daily_n1', 'daily_n2', 'weekly_n1', 'weekly_n2', 'analysis'):
        bad(['3PP 與 SA Care 合併', '3PP與SA Care合併', '3PP（含 SA Care）'], '3PP 不含 Care')
    for m in re.finditer(r'(區內|全區|整區)?(最高|最低|唯一|第一|墊底)', text):
        out.append(('⚠️', '排序／唯一的說法', f'「{text[max(0, m.start() - 10):m.end() + 4]}」要用事實檔排序確認過'))
    if profile in ('daily_n1', 'daily_n2'):
        n = len(re.sub(r'\s', '', text))
        if n < 200:
            out.append(('⚠️', '長度（要填滿 D43，約 220 字以上）', f'全文只有 {n} 字'))
        for p in paras if profile == 'daily_n1' else []:      # Yuni（北二）本來就句號結尾
            if p.rstrip().endswith('。'):
                out.append(('⚠️', '句尾通常無句號', f'{p[-15:]}'))
            if '。' in p.rstrip()[:-1]:
                out.append(('⚠️', '整段逗號串接', '段落中間有句號，Bob 習慣整段逗號串接'))
        if re.search(r'月底前.{0,6}剩[餘下]?\s*\d+\s*天', text):
            out.append(('⚠️', '收尾別每天都「月底前剩餘 X 天」', '10/28 回饋：講法太像，看近幾天有沒有用過'))
        if re.search(r'每[天日]需?(完成|銷售|達成)', text):
            out.append(('⚠️', '每天需完成 X 的換算', '先跟過去日均比，超過 1.5 倍就別寫'))
    if profile == 'daily_n1':
        if re.search(r'(總營業額|營業額)[^，。]{0,8}達成率', text) and not re.search(r'本週|前週|上週', text):
            out.append(('❌', '北一不寫月營業額進度', '禮券回補會讓月營業額失真（10/27 指定）'))
    if profile in ('weekly_n1', 'weekly_n2'):
        if re.search(r'AC\+\s*搭售率[^，。]{0,12}合計', text):
            out.append(('⚠️', '保固口徑', 'AC+＋SA Care 要寫「保固合計搭售率」'))
    if profile == 'weekly_n2' and not speech:
        for p in paras:
            n = len(re.sub(r'\s', '', p))
            if n > 110:
                out.append(('❌' if n > 130 else '⚠️', '北二每格約 110 字內（Yuni 定稿多在 110–128）',
                            f'{n} 字：{p[:20]}…'))
            if not p.rstrip().endswith('。'):
                out.append(('❌', '北二每格結尾加「。」', p[-15:]))
    if profile == 'weekly_n1' and not speech:
        for p in paras:
            if p.rstrip().endswith('。'):
                out.append(('⚠️', '北一週報評語句尾通常無句號', p[-15:]))
    return out


# ---------------------------------------------------------------- 核對單
def report(text, facts, profile, stores, extra=None, title='核對單', rule_text=None, speech=False):
    nums = check_numbers(text, facts, stores)
    rules = rule_results(rule_text if rule_text is not None else text, profile, speech=speech)
    extra = extra or []
    n_bad = sum(r['status'] == '❌' for r in nums) + sum(s == '❌' for s, _, _ in rules + extra)
    n_warn = sum(r['status'] == '⚠️' for r in nums) + sum(s == '⚠️' for s, _, _ in rules + extra)
    lines = [f'## {title}', '',
             f'數字 {len(nums)} 個：✅ {sum(r["status"] == "✅" for r in nums)}｜⚠️ {sum(r["status"] == "⚠️" for r in nums)}'
             f'｜❌ {sum(r["status"] == "❌" for r in nums)}｜🎯 {sum(r["status"] == "🎯" for r in nums)}；規則與資料檢查 ❌ {sum(s == "❌" for s, _, _ in rules + extra)}'
             f'｜⚠️ {sum(s == "⚠️" for s, _, _ in rules + extra)}', '']
    for s, name, why in extra + rules:
        lines.append(f'- {s} {name}：{why}')
    if extra or rules:
        lines.append('')
    lines.append('| | 數字 | 上下文 | 對到 |')
    lines.append('|---|---|---|---|')
    for r in nums:
        lines.append(f"| {r['status']} | {r['tok']['text']} | {r['store'] or '區'}／{r['metric'] or '-'} | {r['why']} |")
    verdict = '❌ 不能交付' if n_bad else ('⚠️ 可交付，但 ⚠️ 項目要逐一確認' if n_warn else '✅ 全部通過')
    lines += ['', f'**結論：{verdict}**']
    return '\n'.join(lines), n_bad, n_warn, nums, rules
