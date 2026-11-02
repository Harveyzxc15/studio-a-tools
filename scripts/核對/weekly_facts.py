# -*- coding: utf-8 -*-
"""週報評語事實檔：把 build_deck.py 產的 評語素材.md 每個數字轉成事實（門市、列、欄、群組）。

評語素材是週報評語唯一的當期數字來源（評語寫作規則.md），所以評語裡的數字一定要對得回這裡。
用法：python3 weekly_facts.py 27Q1W04 [--region 北二區] [--out 路徑]
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from factcheck import fact, save_facts, extract_numbers, N1_STORES, N2_STORES, STORE_ALIAS
from daily_facts import OUT_BASE

DECK = {'北一區': os.path.expanduser('~/工具中心/輸出/北一區週報簡報'),
        '北二區': os.path.expanduser('~/工具中心/輸出/北二區週報簡報')}
NUM = re.compile(r'^-?\d[\d,]*(\.\d+)?%?$')
PROD = re.compile(r'(CPU|Mac|Watch|iPad|iPhone|AirPods|Airpods)')
PERIOD = re.compile(r'^\d{1,4}[ /]?\d{0,2}/\d{2}~')


def with_product(cells):
    """表頭裡產品名之後的欄位都屬於該產品（CPU | Mac ACPP+ | ACPP+搭售率 … | Watch | …）；
    開頭的期間（10/11~10/17）拿掉，讓本週表、上週表同一欄同群組。"""
    out, cur = [], ''
    for i, c in enumerate(cells):
        if i == 0 and PERIOD.match(c):
            out.append('')
            continue
        m = PROD.search(c)
        cur = m.group(1) if m else cur
        out.append(c if (not cur or cur in c) else f'{cur} {c}')
    return out


def store_in(text, stores):
    for s in sorted(list(STORE_ALIAS) + stores, key=len, reverse=True):
        if s in text:
            return STORE_ALIAS.get(s, s)
    if re.search(r'Total|合計|北一區|北二區|門市加總', text):
        return ''
    return None


def to_val(cell, header):
    pct = cell.endswith('%')
    v = float(cell.rstrip('%').replace(',', ''))
    if pct:
        return v / 100, 'rate'
    # 「週對比 本週／上週／差異」是金額欄：「週對比」的「比」不算比率
    if re.search(r'率|比|佔比|％|%', header.replace('週對比', '')) or ('.' in cell and abs(v) <= 1.5):
        return v, 'rate'
    return v, ('amt' if abs(v) >= 1000 else None)


def parse(md, stores):
    facts = parse_raw(md, stores)
    # 標題「本週 10/11~10/17，上週 10/04~10/10」→ 標籤裡的日期換成本週／上週，核對時才能比對期間
    m = re.search(r'本週\s*(\d{1,2}/\d{2})~(\d{1,2}/\d{2})，上週\s*(\d{1,2}/\d{2})~(\d{1,2}/\d{2})', md)
    if m:
        this, prev = f'{m.group(1)}~{m.group(2)}', f'{m.group(3)}~{m.group(4)}'
        for f in facts:
            f['l'] = f['l'].replace(this, '本週').replace(prev, '上週')
    return facts


def parse_raw(md, stores):
    facts = []
    section, sec_store, hdr, top = '', None, [], []
    prev_header = False
    for raw in md.split('\n'):
        line = raw.strip()
        if line.startswith('## '):
            section, sec_store, hdr, top, prev_header = line[3:], store_in(line[3:], stores), [], [], False
            if sec_store == '':
                sec_store = None
            continue
        if line.startswith('### '):
            sec_store = store_in(line[4:], stores)
            continue
        if not line:
            continue
        if '|' not in line or section.startswith('上週簡報的評語'):
            # 文字行：上週評語或說明文字，數字照單位存（「上週 ARpedia 僅 1 套」要能被引用）
            who = store_in(line, stores) if store_in(line, stores) is not None else (sec_store or '')
            for t in extract_numbers(line):
                v = t['x'] / 100 if t['unit'] in ('%', '％') else t['x'] * (1e4 if t['unit'] == '萬' else 1)
                u = 'rate' if t['unit'] in ('%', '％') else 'amt' if t['unit'] in ('萬', '元') or t['comma'] else 'cnt'
                start = max([line.rfind(c, 0, t['pos']) for c in '。'] + [-1]) + 1
                uc = f'〔{t["unit"]}〕' if t['unit'] else ''
                facts.append(fact(v, u, who, f'{who or "區"} {section[:8]}：{line[max(start, t["pos"] - 40):t["end"]]}{uc}',
                                  src='評語素材文字'))
            continue
        cells = [c.strip() for c in line.split('|')]
        n_num = sum(bool(NUM.match(c) or re.match(r'^-?\d+(/[-\d.]+)+$', c)) for c in cells if c)
        n_cell = sum(1 for c in cells if c)
        # 「士林 4 | 微風 3」「士林 63/108 | 微風 31/62」：店名＋數字在同一格
        pairs = [re.match(r'^(\S+?)\s+([-\d.,/]+)$', c) for c in cells]
        if n_cell and all(p and store_in(p.group(1), stores) is not None for p in pairs if p) and sum(map(bool, pairs)) >= 2:
            label = re.sub(r'[：:].*$', '', line) if '：' in line else section
            for p in pairs:
                if not p:
                    continue
                who = store_in(p.group(1), stores)
                for i, part in enumerate(p.group(2).split('/')):
                    if NUM.match(part):
                        facts.append(fact(float(part.replace(',', '')), None, who, f'{who} {label[:16]} #{i + 1}',
                                          g=f'{section}|{label}|{i}', src='評語素材'))
            continue
        if n_num <= n_cell * 0.3:              # 表頭
            if prev_header:
                ff, out = '', []
                for i, c in enumerate(cells):
                    t = hdr[i] if i < len(hdr) else ''
                    ff = t or ff
                    out.append(ff)
                top = out
                hdr = [(top[i] + ' ' + c).strip() if i < len(top) else c for i, c in enumerate(cells)]
            else:
                hdr, top = with_product(cells), []
            prev_header = True
            continue
        prev_header = False
        row_label, row_store = '', None
        for i, c in enumerate(cells):
            if not c:
                continue
            h = hdr[i] if i < len(hdr) else ''
            t = top[i] if i < len(top) else ''
            if not (NUM.match(c) or re.match(r'^-?[\d.]+(/[-\d.]+)+$', c)):
                s = store_in(c, stores)
                if s is not None:
                    row_store = s
                else:
                    row_label = c
                continue
            col_store = store_in(h, stores)
            who = row_store if row_store is not None else col_store if col_store is not None else (sec_store or '')
            item = row_label if row_store is not None or col_store is not None else row_label
            title = hdr[0] if hdr and hdr[0] and not NUM.match(hdr[0]) else ''
            col = re.sub(r'\d{1,4}/\d{2}~\d{1,2}/?\d{2}|20\d\d|本週|上週|對照', '', h).strip()
            grp_row = f'{section}|{title}|{who}|{item}|{col}'   # 同表同店同列同欄（去掉期間）：本週 vs 上週
            grp_col = f'{section}|{title}|{item}|{h}'        # 同列同欄跨店：只拿來算店 vs 區
            if '/' in c:
                parts = c.split('/')
                names = (h.split()[-1].split('/') if h else [])
                vals = []
                for k, part in enumerate(parts):
                    if NUM.match(part):
                        nm = names[k] if k < len(names) else f'#{k + 1}'
                        vals.append(float(part))
                        facts.append(fact(float(part), 'cnt', who, f'{who or "區"} {item} {h} {nm}',
                                          g=f'{section}|{who}|{item}|{nm}', src='評語素材'))
                if len(vals) == 3 and vals[0] + vals[2]:
                    facts.append(fact(vals[2] / (vals[0] + vals[2]), 'rate', who,
                                      f'{who or "區"} {item} {h} 零搭售佔比', src='評語素材衍生'))
                continue
            v, u = to_val(c, h)
            label = f'{who or "區"} {item} {h} [{title}]'.strip()
            facts.append(fact(v, u, who, label, g=grp_row, src='評語素材'))
            facts.append(fact(v, u, who, label, g=grp_col, src='評語素材'))
    return facts


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    label = args[0]
    region = '北二區' if '--region' in sys.argv and '北二' in sys.argv[sys.argv.index('--region') + 1] else '北一區'
    src = os.path.join(DECK[region], label, '評語素材.md')
    out = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else \
        os.path.join(OUT_BASE, f'週報_{region}_{label}', 'facts.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(src, encoding='utf-8') as f:
        facts = parse(f.read(), N2_STORES if region == '北二區' else N1_STORES)
    save_facts(out, facts, {'kind': 'weekly', 'region': region, 'label': label, 'source': src, 'checks': []})
    print(f'事實 {len(facts)} 筆 → {out}')


if __name__ == '__main__':
    main()
