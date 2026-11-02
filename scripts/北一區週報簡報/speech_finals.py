"""取回使用者實際上台講的週報講稿（定稿），和我寫的初稿比對。

    python3 speech_finals.py 27Q1W04                 # 北一：從「週三會議」Google 文件抓 27Q1W04 分頁
    python3 speech_finals.py 27Q1W04 --region 北二區  # 北二目前沒有來源，會提示

來源設定在 <輸出>/講稿定稿來源.json（{"北一區": {"doc_id": ...}}）：文件是「知道連結就能讀」，
doc_id 等於公開連結，所以放在不 commit 的輸出資料夾。
存成 <輸出>/<週次>/講稿_定稿.md，並在 講稿修改紀錄.md 記一筆相似度；改了什麼、歸納成什麼規則，由寫講稿的人讀兩份後補寫。
build_deck（北一）每週會自動抓上一週的。
"""
import argparse
import json
import re
import subprocess
from difflib import SequenceMatcher

import build_deck as b

LOG = b.HERE / '講稿修改紀錄.md'


def _get(url):
    # 用 curl：python.org 版 Python 沒裝 SSL 憑證（urllib 會 CERTIFICATE_VERIFY_FAILED）
    r = subprocess.run(['curl', '-sSLf', '-A', 'Mozilla/5.0', url], capture_output=True, timeout=120)
    if r.returncode != 0:
        raise SystemExit(f'❌ 讀取 Google 文件失敗：{r.stderr.decode()[:200]}')
    return r.stdout.decode('utf-8', 'ignore')


def tabs(doc_id):
    """{分頁名: 分頁 id}；從文件網頁內嵌的資料解析（不需登入，文件要是「知道連結就能讀」）。"""
    html = _get(f'https://docs.google.com/document/d/{doc_id}/edit')
    out = {}
    for tid, name in re.findall(r'"d":\["(t\.[a-z0-9]+)",\[1,"([^"]+)"\]', html):
        out.setdefault(name.strip(), tid)
    return out


def _versions(text):
    """講稿文字 → (完整版, 一分半版)；以「一分半版」那一行切開。"""
    text = re.sub(r'^﻿?\s*\d\dQ\dW\d\d\s*', '', text.strip())
    text = re.sub(r'^# .*$', '', text, flags=re.M)            # 只去掉最上面的大標題（## 完整版／一分半版要留著切）
    parts = re.split(r'^\s*#*\s*一分半版\s*$', text, maxsplit=1, flags=re.M)
    full = re.sub(r'^\s*#*\s*完整版\s*$', '', parts[0], flags=re.M)
    short = parts[1] if len(parts) > 1 else ''
    short = short.split('________________')[0]          # 分頁後面常接別的筆記（客訴說明等）
    clean = lambda s: re.sub(r'\s+', '', s)
    return clean(full), clean(short)


def fetch(label):
    cfg_p = b.OUT_BASE / '講稿定稿來源.json'
    cfg = json.loads(cfg_p.read_text(encoding='utf-8')).get(b.REGION) if cfg_p.exists() else None
    if not cfg:
        raise SystemExit(f'{b.REGION}沒有講稿定稿來源（{cfg_p}）；請使用者把實際講的版本貼回來，存成 {label}/講稿_定稿.md')
    t = tabs(cfg['doc_id'])
    tid = t.get(label)
    if not tid:
        raise SystemExit(f'週三會議文件還沒有 {label} 分頁（最近：{list(t)[-3:]}）')
    text = _get(f"https://docs.google.com/document/d/{cfg['doc_id']}/export?format=txt&tab={tid}")
    work = b.OUT_BASE / label
    work.mkdir(parents=True, exist_ok=True)
    out = work / '講稿_定稿.md'
    out.write_text(text.lstrip('﻿'), encoding='utf-8')
    draft = work / '講稿.md'
    line = f'- {b.REGION} {label}：定稿已取回（{out}）'
    if draft.exists():
        d, f = _versions(draft.read_text(encoding='utf-8')), _versions(text)
        sims = [f'{n} {SequenceMatcher(None, x, y).ratio():.0%}' for n, x, y in zip(('完整版', '一分半版'), d, f) if x and y]
        line += '，與初稿相似度 ' + '、'.join(sims) + '；改了什麼請讀兩份後歸納進上面的寫法規則'
    else:
        line += '，這週沒有我寫的初稿（使用者自己寫的，可當寫法範例）'
    log = LOG.read_text(encoding='utf-8') if LOG.exists() else ''
    if f'{b.REGION} {label}：' not in log:
        LOG.write_text(log.rstrip('\n') + '\n' + line + '\n', encoding='utf-8')
    b.log(line)
    return out


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('label', help='例：27Q1W04')
    ap.add_argument('--region', default='北一區', choices=list(b.REGIONS))
    a = ap.parse_args()
    b.configure(a.region)
    fetch(a.label)
