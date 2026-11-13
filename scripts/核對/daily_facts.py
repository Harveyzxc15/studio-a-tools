# -*- coding: utf-8 -*-
"""日報事實檔：讀各店門市日報附件（0XX_店名日報_YYYYMMDD.xlsx，data_only 有值），
把每一列的 今日／本月累積／本月目標／上月同期、達成率、與上月比 都寫進事實檔，並算區合計與各種率。

用法：python3 daily_facts.py 20261031 [--region 北二區] [--out 路徑]
預設輸出 ~/工具中心/輸出/核對/日報_<區>_<日期>/facts.json
"""
import os
import re
import subprocess
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from factcheck import fact, save_facts

# 自動盲測在 mini 跑：讀信件備份另存的日報附件（launchd 讀不到 ~/Library/Mail）
MAIL = os.environ.get('DAILY_XLSX_DIR') or os.path.expanduser('~/Library/Mail')
OUT_BASE = os.path.expanduser('~/工具中心/輸出/核對')
STORES = {'北一區': {'004': '士林', '005': '微風', '024': '美麗華', '046': '阿波羅', '054': '高島屋', '057': '羅東'},
          '北二區': {'009': '永和', '025': '板橋誠品', '050': '西門', '055': '花蓮', '063': '板橋遠百',
                  '064': '新莊宏匯', '068': '裕隆城'}}
PERIODS = {3: '今日', 4: '月累', 5: '目標', 7: '上月同期'}
# 列號 → (標籤, 單位)。率的列另外由合計重算
ROWS = {7: ('總營業額', 'amt'), 8: ('零售營業額', 'amt'), 9: ('訂單出貨', 'amt'), 11: ('訂金', 'amt'),
        13: ('Apple主機營業額', 'amt'), 14: ('Apple配件營業額', 'amt'), 15: ('3PP營業額', 'amt'),
        16: ('3PP搭售率', 'rate'), 17: ('保固營業額', 'amt'), 18: ('保固金額搭售率', 'rate'),
        19: ('總毛利', 'amt'), 20: ('Apple毛利', 'amt'), 21: ('3PP毛利', 'amt'), 22: ('保固毛利', 'amt'),
        23: ('CPU台數', 'cnt'), 24: ('CPU保固套數', 'cnt'), 25: ('CPU保固搭售率', 'rate'),
        26: ('iPad台數', 'cnt'), 27: ('iPad保固套數', 'cnt'), 28: ('iPad保固搭售率', 'rate'),
        29: ('iPhone台數', 'cnt'), 30: ('iPhone保固套數', 'cnt'), 31: ('iPhone保固搭售率', 'rate'),
        32: ('Watch台數', 'cnt'), 33: ('Watch保固套數', 'cnt'), 34: ('Watch保固搭售率', 'rate'),
        35: ('AirPods台數', 'cnt'), 36: ('AirPods保固套數', 'cnt'), 37: ('AirPods保固搭售率', 'rate'),
        38: ('人均產值', 'amt'), 39: ('來客數', 'cnt'), 40: ('成交筆數', 'cnt'), 41: ('提袋率', 'rate'),
        42: ('舊換新台數', 'cnt')}
# 率＝分子列÷分母列（區合計與重算用）
RATES = {16: (15, 7), 18: (17, 7), 25: (24, 23), 28: (27, 26), 31: (30, 29), 34: (33, 32), 37: (36, 35),
         41: (40, 39)}


def find_files(day, codes):
    out = subprocess.run(['find', MAIL, '-name', f'*日報_{day}*.xlsx'], capture_output=True, text=True).stdout
    paths = [p for p in out.split('\n') if re.search(rf'/({"|".join(codes)})_[^/]*日報_{day}(-\d+)?\.xlsx$', p)]
    best = {}
    for p in sorted(paths, key=os.path.getmtime, reverse=True):
        c = re.search(r'/(\d{3})_', p).group(1)
        best.setdefault(c, p)
    return best


def num(v):
    return v if isinstance(v, (int, float)) else 0


def build(day, region, allow_missing=False):
    """allow_missing：缺店照算（⚠️ 不判 ❌），但不出區合計／店佔區（缺店加總會偏低）。"""
    stores = STORES[region]
    files = find_files(day, list(stores))
    facts, checks, raw = [], [], {}
    missing = [stores[c] for c in stores if c not in files]
    if missing:
        checks.append(('⚠️' if allow_missing else '❌', '資料完整',
                       f'{day} 缺門市日報：{"、".join(missing)}（信還沒到齊，不要用其他店推估）'))

    pace = None
    for code, path in files.items():
        name = stores[code]
        ws = openpyxl.load_workbook(path, data_only=True)['日報']
        d = ws['C4'].value
        if not isinstance(ws['D7'].value, (int, float)):
            checks.append(('❌', '資料完整', f'{name} 日報附件只有公式沒有值（總營業額月累是空的），'
                                          f'區合計會漏掉這店；改讀已開啟的 Excel 或等重寄：{path}'))
            continue
        if ws['C3'].value and name[:2] not in str(ws['C3'].value):
            checks.append(('❌', '店名', f'{path} 的店名是 {ws["C3"].value}'))
        raw[name] = {r: {c: num(ws.cell(r, c).value) for c in PERIODS} for r in ROWS}
        pace = num(ws['H5'].value) or pace
        for r, (label, u) in ROWS.items():
            for c, per in PERIODS.items():
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)):
                    facts.append(fact(v, u, name, f'{name} {label} {per}', g=f'{label}|{per}', src='門市日報', p=per))
                    facts.append(fact(v, u, name, f'{name} {label} {per}', g=f'{name}|{label}', src='門市日報', p=per))
            for c, per in ((6, '達成率'), (8, '與上月同期比')):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)):
                    if per == '與上月同期比' and u == 'rate':
                        facts.append(fact(abs(v), 'rate', name, f'{name} {label} 比上月(百分點)', src='門市日報'))
                    else:
                        facts.append(fact(abs(v) if per != '達成率' else v, 'rate', name, f'{name} {label} {per}',
                                          g=f'{label}|{per}', src='門市日報'))
        if d and int(d) != int(day[6:]):
            checks.append(('❌', '日期', f'{name} 檔案日期 {d} 與 {day} 不符'))
    # 區合計（缺店時不出，加總會偏低）
    if raw and not missing:
        tot = {r: {c: sum(raw[s][r][c] for s in raw) for c in PERIODS} for r in ROWS}
        for r, (a, b) in RATES.items():
            for c in PERIODS:
                tot[r][c] = tot[a][c] / tot[b][c] if tot[b][c] else 0
        for r, (label, u) in ROWS.items():
            if r == 38:
                continue
            for c, per in PERIODS.items():
                facts.append(fact(tot[r][c], u, '', f'{region} {label} {per}', g=f'{label}|{per}', src='六店加總', p=per))
                facts.append(fact(tot[r][c], u, '', f'{region} {label} {per}', g=f'|{label}', src='六店加總', p=per))
            if tot[r][5]:
                facts.append(fact(tot[r][4] / tot[r][5], 'rate', '', f'{region} {label} 達成率',
                                  g=f'{label}|達成率', src='六店加總'))
        # 店佔區：例「士林佔全區 iPhone 32%」
        for s in raw:
            for r, (label, u) in ROWS.items():
                for c, per in PERIODS.items():
                    if (r in (23, 26, 29, 32, 35, 7, 15, 17)) and tot[r][c]:
                        facts.append(fact(raw[s][r][c] / tot[r][c], 'rate', s, f'{s} {label} {per} 佔區',
                                          src='衍生'))
    if not missing and not any(c[0] == '❌' and c[1] == '資料完整' for c in checks):
        checks.append(('✅', '資料完整', f'{day} {len(stores)} 店門市日報都在且有值'))
    if pace:
        facts.append(fact(pace, 'rate', '', '本月應達成率', src='門市日報 H5'))
    d = int(day[6:])
    facts.append(fact(d, 'cnt', '', '本月已過天數', src='日期'))
    import datetime as _dt
    dt = _dt.date(int(day[:4]), int(day[4:6]), d)
    n_year = 366 if dt.year % 4 == 0 else 365
    facts.append(fact(dt.timetuple().tm_yday / n_year, 'rate', '', '年度應達成率', src='日期（今年第幾天÷全年天數）'))
    return facts, checks


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    day = args[0]
    region = '北二區' if '--region' in sys.argv and '北二' in sys.argv[sys.argv.index('--region') + 1] else '北一區'
    out = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else \
        os.path.join(OUT_BASE, f'日報_{region}_{day}', 'facts.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    facts, checks = build(day, region)
    save_facts(out, facts, {'kind': 'daily', 'region': region, 'day': day, 'checks': checks})
    for c in checks:
        print(*c)
    print(f'事實 {len(facts)} 筆 → {out}')


if __name__ == '__main__':
    main()
