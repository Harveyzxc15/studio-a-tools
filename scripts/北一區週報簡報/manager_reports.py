# -*- coding: utf-8 -*-
"""店長週報（Google Drive「店長週報」）→ 只取店長寫的評語文字，當寫週報評語的背景。

  python3 manager_reports.py 27Q1W06 [--region 北二區]

- 資料夾是連結共用，未登入可讀：店長週報/<區>/<20YY>/Q#/W#/2027Q1W6<店>門市X級.pdf
- 下載到 ~/工具中心/輸出/店長週報評語/<區>/<週次>/pdf/，用 pypdf 抽「本週門市狀況／執行進度與改善狀況（上週成效、本週預計）／
  CPU 檢討報告」這幾頁的文字，存 <區>_<週次>.md／.json
- 標出會讓數字失真的背景（大量採購、盤商、送禮、外國客、缺貨、家人阻止／預算、教育價空機），附原句與品類
- 評語素材加「## 店長週報」一節：各店背景＋上週店長說要做什麼 → 本週回報成效 → 本週預計
- **數字一律以週報工具為準**（使用者 2026-11-09 指定），店長週報只當背景，評語不引用店長給的數字
使用者說「做週報」時由 build_deck 自動抓（不排程，使用者 2026-11-09 指定）；店長還沒傳齊時再跑這支補齊（會一併重寫追蹤帳本／選題矩陣）。
"""
import argparse
import json
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / '核對'))
import topics as T  # noqa: E402
import build_deck as B  # noqa: E402

ROOT_FOLDER = '1TYz0ALbQEvtzad-nG31dEgrCGn00_4jk'      # 店長週報
OUT = Path.home() / '工具中心' / '輸出' / '店長週報評語'
FLAGS = [   # (類型, 正則)
    ('大量採購', r'大量採購|大量購|盤商|公司用機|公司採購|採購\s*\d|團購'),
    ('送禮', r'送禮|送人|買送|送女友|送男友|禮物'),
    ('外國客', r'外國|外籍|退稅|觀光客'),
    ('缺貨', r'缺貨|沒貨|無現貨|沒有現貨|調不到|買不到|欠貨|等待到貨|等到貨|沒有其他保護殼'),
    ('家人／預算', r'家人|預算|現金'),
    ('教育價空機', r'EDU|教育價.*空機|空機'),
]


def norm(s):
    """部首相容字（⼈⾨⽇）轉一般字；其他字不動（全形標點保留）。"""
    return ''.join(unicodedata.normalize('NFKC', ch) if 0x2E80 <= ord(ch) <= 0x2FDF else ch for ch in s)


# ───────────── Drive ─────────────
def _curl(url, out=None):
    cmd = ['curl', '-sL', '--max-time', '120', url] + (['-o', str(out)] if out else [])
    r = subprocess.run(cmd, capture_output=True, timeout=150)
    if r.returncode != 0:
        raise RuntimeError(f'下載失敗：{url}')
    return r.stdout.decode('utf-8', 'ignore') if not out else None


def list_folder(fid):
    """[(名稱, id, 是否資料夾)]"""
    h = _curl(f'https://drive.google.com/embeddedfolderview?id={fid}')
    out = []
    for m in re.finditer(r'<a href="https://drive\.google\.com/(drive/folders|file/d)/([\w-]+)[^"]*"[^>]*>.*?'
                         r'<div class="flip-entry-title">([^<]+)</div>', h, re.S):
        out.append((m[3].strip(), m[2], m[1] == 'drive/folders'))
    return out


def week_folder(region, label):
    """店長週報/<區>/20YY/Q#/W# 的資料夾 id；找不到回 None。"""
    m = re.fullmatch(r'(\d\d)Q(\d)W(\d\d)', label)
    path = [region, f'20{m[1]}', f'Q{m[2]}', f'W{int(m[3])}']
    fid = ROOT_FOLDER
    for name in path:
        hit = next((i for n, i, d in list_folder(fid) if d and n.strip() in (name, f'W{m[3]}')), None)
        if not hit:
            return None
        fid = hit
    return fid


def store_of(fname):
    base = re.sub(r'^\d{4}Q\dW\d+', '', fname)
    base = re.sub(r'\s*[ABC]級.*$|\.pdf$', '', base).replace('門市', '').strip()
    for _code, name, full in B.STORES:              # 「新西門」→西門、「宏匯」→新莊宏匯
        if name in base or base in name or base in full:
            return name
    return B.store_key(base) or base


def fetch(region, label):
    """下載本週所有店長週報 → {店: pdf 路徑}；資料夾還沒建好回 {}。"""
    fid = week_folder(region, label)
    if not fid:
        return {}
    d = OUT / region / label / 'pdf'
    d.mkdir(parents=True, exist_ok=True)
    got = {}
    for name, i, is_dir in list_folder(fid):
        if is_dir or not name.lower().endswith('.pdf'):
            continue
        p = d / name
        if not p.exists() or p.stat().st_size < 1000:
            _curl(f'https://drive.google.com/uc?export=download&id={i}', p)
        if p.read_bytes()[:4] == b'%PDF':
            got[store_of(name)] = p
    return got


# ───────────── 抽評語 ─────────────
def _between(t, start, ends):
    i = t.find(start)
    if i < 0:
        return ''
    j = min([t.find(e, i + len(start)) for e in ends if t.find(e, i + len(start)) > 0] or [len(t)])
    return t[i + len(start):j].strip(' ：:\n')


def _clean(s):
    s = re.sub(r'\(請[^)]*\)|（請[^）]*）', '', s)          # 範本裡的填寫說明
    s = re.sub(r'[ \t]*\n[ \t]*', '\n', s)
    s = re.sub(r'\n(?=[^\n\d一二三四五六七八九十*\-•■]{1,6}\n)', '', s)   # 被換行切斷的短句接回去
    return re.sub(r'\n{2,}', '\n', s).strip()


def extract(pdf):
    from pypdf import PdfReader
    import logging
    logging.disable(logging.CRITICAL)
    pages = [norm(p.extract_text() or '') for p in PdfReader(str(pdf)).pages]
    res = {'status': '', 'done': '', 'plan': '', 'cpu_review': ''}
    for t in pages:
        if '本週門市狀況' in t and not res['status']:
            res['status'] = _clean(_between(t, '本週門市狀況', ['執行進度與改善狀況', '修繕']))
            prog = _between(t, '執行進度與改善狀況', ['修繕項目', '修繕/', '本月尚未完成'])
            res['done'] = ''
            for head in ('成效', '執行成果', '的執行進度', '執行進度'):
                res['done'] = _clean(_between(prog, head, ['本週預計執行', '本週預計']))
                if res['done']:
                    break
            res['plan'] = _clean(_between(prog, '本週預計執行的重點項目', ['修繕']) or _between(prog, '本週預計', ['修繕'])
                                 or _between(prog, '努力方向', ['修繕']) or _between(prog, '如何改善', ['本週預計']))
            if not (res['done'] or res['plan']):
                res['plan'] = _clean(prog)
        if '檢討報告' in t and not res['cpu_review']:
            res['cpu_review'] = _clean(t.split('檢討報告', 1)[1])
    res['flags'] = flags_of('\n'.join(res[k] for k in ('status', 'done', 'cpu_review')))
    return res


def flags_of(text):
    out = []
    for sent in re.split(r'[。\n；]', text):
        for kind, pat in FLAGS:
            if re.search(pat, sent):
                s = sent.strip(' ，,')
                devs = [d for d, p in T._DEV if re.search(p, s)]
                out.append({'kind': kind, 'devices': devs, 'text': s[:90]})
    return out


# ───────────── 產出 ─────────────
def run(label, region, work=None, refresh=True):
    """抓＋抽本週（與上週，供「上週說要做」）→ json／md；回傳 {店: 抽出內容}。"""
    B.configure(region)
    data = {}
    for lab in (label, prev_label(label)):
        if not lab:
            continue
        jp = OUT / f'{region}_{lab}.json'
        if lab != label and jp.exists():
            data[lab] = json.loads(jp.read_text(encoding='utf-8'))
            continue
        pdfs = fetch(region, lab) if (refresh or not jp.exists()) else {}
        if pdfs:
            d = {s: extract(p) for s, p in pdfs.items()}
            jp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding='utf-8')
            (OUT / f'{region}_{lab}.md').write_text(to_md(region, lab, d), encoding='utf-8')
        if jp.exists():
            data[lab] = json.loads(jp.read_text(encoding='utf-8'))
    cur, prev = data.get(label, {}), data.get(prev_label(label), {})
    work = Path(work) if work else B.OUT_BASE / label
    if (work / '評語素材.md').exists():
        _write_section(work / '評語素材.md', material_md(cur, prev))
    return cur


def prev_label(label):
    m = re.fullmatch(r'(\d\d)Q(\d)W(\d\d)', label)
    fy, q, w = int(m[1]), int(m[2]), int(m[3])
    if w > 1:
        return f'{fy}Q{q}W{w - 1:02d}'
    return None                                         # 季初第一週：上一季最後一週的週數不固定，不抓


def to_md(region, label, d):
    md = [f'# {region} 店長週報評語 {label}', '', '來源：Google Drive「店長週報」，只取店長寫的文字；數字以週報工具為準。', '']
    for s in B.NAMES:
        if s not in d:
            md += [f'## {s}', '（尚未上傳）', '']
            continue
        x = d[s]
        md += [f'## {s}', '### 本週門市狀況', x['status'] or '（空白）', '### 上週重點項目成效', x['done'] or '（空白）',
               '### 本週預計執行', x['plan'] or '（空白）', '### CPU 檢討報告', x['cpu_review'] or '（空白）', '']
    return '\n'.join(md)


def material_md(cur, prev):
    md = ['## 店長週報（只當背景，數字以本素材為準）', '',
          '店長在週報裡寫的說明。「數字背景」會讓零搭售、搭售率失真（大量採購、送禮、外國客、缺貨、家人阻止…），'
          '寫評語時要考慮進去，可寫「含一筆大量採購」，但不要引用店長給的數字。「上週說要做 → 本週回報」可用來追店長自己的承諾。', '']
    if not cur:
        md.append('本週店長週報還沒上傳（店長多半週一晚上才傳齊；之後跑 `python3 manager_reports.py <週次>` 補上）。')
        return '\n'.join(md) + '\n'
    missing = [s for s in B.NAMES if s not in cur]
    if missing:
        md.append(f'尚未上傳：{"、".join(missing)}\n')
    for s in B.NAMES:
        if s not in cur:
            continue
        x = cur[s]
        md.append(f'### {s}')
        fl = x.get('flags', [])
        if fl:
            md.append('- 數字背景：' + '；'.join(f'【{f["kind"]}{"·" + "/".join(f["devices"]) if f["devices"] else ""}】{f["text"]}' for f in fl[:8]))
        if s in prev and prev[s].get('plan'):
            md.append(f'- 上週店長說要做：{_one(prev[s]["plan"])}')
        if x.get('done'):
            md.append(f'- 本週店長回報成效：{_one(x["done"])}')
        elif s in prev:
            md.append('- 本週店長回報成效：（空白）')
        if x.get('plan'):
            md.append(f'- 本週店長預計：{_one(x["plan"])}')
        md.append(f'- 本週門市狀況（摘）：{_one(x["status"], 260)}')
        md.append('')
    return '\n'.join(md) + '\n'


def _one(s, n=200):
    s = re.sub(r'\s*\n\s*', ' / ', s or '').strip(' /')
    return s if len(s) <= n else s[:n] + '…'


def _write_section(mat_path, body):
    txt = mat_path.read_text(encoding='utf-8')
    txt = re.sub(r'\n## 店長週報.*?(?=\n## |\Z)', '', txt, flags=re.S)
    head, sep, rest = txt.partition('\n## 追蹤帳本')
    if not sep:
        head, sep, rest = txt.partition('\n## 上週快照')
    mat_path.write_text(head.rstrip() + '\n\n' + body + sep + rest, encoding='utf-8')


def load(region, label):
    p = OUT / f'{region}_{label}.json'
    return json.loads(p.read_text(encoding='utf-8')) if p.exists() else {}


def auto():
    """不帶週次時：兩區各找最近 4 天內 build_deck 做過的週次，補抓店長週報並重寫素材，最後跳 macOS 通知。"""
    import time
    import tracker
    msgs = []
    for region in ('北一區', '北二區'):
        B.configure(region)
        weeks = sorted((d for d in B.OUT_BASE.iterdir() if d.is_dir() and re.fullmatch(r'\d\dQ\dW\d\d', d.name)
                        and (d / '評語素材.md').exists()
                        and time.time() - (d / '評語素材.md').stat().st_mtime < 4 * 86400), key=lambda d: d.name)
        if not weeks:
            print(f'{region}：最近 4 天沒有做週報，略過')
            continue
        label = weeks[-1].name
        try:
            cur = run(label, region)
            tracker.run(label, region)
            B.configure(region)
            miss = [s for s in B.NAMES if s not in cur]
            msgs.append(f'{region[:2]} {len(cur)}/{len(B.NAMES)}' + (f'（缺{"、".join(miss)}）' if miss else ''))
            print(f'{time.strftime("%m/%d %H:%M")} {region} {label}：已讀 {len(cur)} 店；尚未上傳：{"、".join(miss) or "無"}')
        except Exception as e:
            msgs.append(f'{region[:2]} 失敗')
            print(f'{time.strftime("%m/%d %H:%M")} {region} {label}：失敗 {e!r}')
    if msgs:
        text = '店長週報已更新到評語素材：' + '、'.join(msgs)
        subprocess.run(['osascript', '-e', f'display notification "{text}" with title "週報評語"'], capture_output=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('label', nargs='?')
    ap.add_argument('--region', default='北一區')
    ap.add_argument('--auto', action='store_true', help='兩區、自動找本週週次')
    a = ap.parse_args()
    if a.auto or not a.label:
        return auto()
    cur = run(a.label, a.region)
    B.configure(a.region)
    print(f'{a.region} {a.label}：已讀 {len(cur)} 店（{"、".join(cur) or "無"}）；'
          f'尚未上傳：{"、".join(s for s in B.NAMES if s not in cur) or "無"}')
    if (B.OUT_BASE / a.label / '評語素材.md').exists():
        import tracker
        tracker.run(a.label, a.region)                 # 選題矩陣要帶店長背景，重寫一次（帳本重跑不會重複）
        print('評語素材已更新：店長週報、追蹤帳本、選題矩陣')


if __name__ == '__main__':
    main()
