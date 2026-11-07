# -*- coding: utf-8 -*-
"""其他區（北一、北二以外）區主管 Daily report 的 D43「本日營運概況」→ 語料/其他區日報.txt（給自動盲測當選題範例）。

來源：筆電 Apple Mail（Envelope Index，歷史完整）＋ mini 信件備份 archive/msgs（每小時備份，只有近期）。
用法：python3 其他區日報_語料.py [起日 YYYY-MM-DD，預設 2026-05-01]
已有的（同日期＋區）不重複加；檔案依日期排序。
"""
import email
import glob
import io
import json
import os
import re
import sqlite3
import sys
from email import policy
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
OUT = HERE / '語料' / '其他區日報.txt'
ARCHIVE = HERE.parent / '信件追蹤' / 'archive'
MD = Path('~/Library/Mail/V10').expanduser()
OWN = re.compile(r'北ㄧ區|北一區|北二區')
SUBJ = re.compile(r'(daily report)', re.I)


def region_day(subject, received=None):
    s = subject.replace(' ', '')
    m = re.search(r'(\d{8})', s)
    reg = re.search(r'([一-鿿]{1,4}區)', s)
    day = m.group(1) if m else None
    if not day:                                   # 「中區 10/14 daily report」
        md = re.search(r'(\d{1,2})/(\d{1,2})', subject)
        if md and received:
            day = f'{received[:4]}{int(md.group(1)):02d}{int(md.group(2)):02d}'
    return day, (reg.group(1) if reg else None)


def d43(blob):
    wb = openpyxl.load_workbook(io.BytesIO(blob), read_only=True, data_only=True)
    ws = wb['日報'] if '日報' in wb.sheetnames else wb.worksheets[0]
    t = [c.value.strip() for r in ws.iter_rows(min_row=40, max_row=60) for c in r
         if isinstance(c.value, str) and len(c.value.strip()) > 30]
    return '\n'.join(t)


def load_existing():
    have, rows = set(), []
    if OUT.exists():
        for blk in OUT.read_text(encoding='utf-8').split('\n## ')[0:]:
            blk = blk.lstrip('# ').strip()
            if not blk:
                continue
            head, _, body = blk.partition('\n')
            k = tuple(head.split()[:2])
            if len(k) == 2:
                have.add(k)
                rows.append((k[0], k[1], body.strip()))
    return have, rows


def from_apple_mail(since):
    db = MD / 'MailData' / 'Envelope Index'
    if not db.exists():
        return
    con = sqlite3.connect(f'file:{db}?immutable=1', uri=True)
    q = """select m.ROWID, s.subject, datetime(m.date_received,'unixepoch','localtime') from messages m
           join subjects s on m.subject=s.ROWID where m.date_received>strftime('%s',?)
           and s.subject like '%aily report%' order by m.date_received"""
    for rid, subj, rec in con.execute(q, (since,)):
        if OWN.search(subj) or not SUBJ.search(subj):
            continue
        day, reg = region_day(subj, rec)
        if not day or not reg:
            continue
        fs = glob.glob(f'{MD}/*/*.mbox/*/Data/*/*/*/Messages/{rid}.emlx') + \
            glob.glob(f'{MD}/*/*.mbox/*/Data/*/*/Messages/{rid}.emlx')
        if not fs:
            continue
        raw = open(fs[0], 'rb').read()
        raw = raw[raw.index(b'\n') + 1:]
        msg = email.message_from_bytes(raw, policy=policy.default)
        for p in msg.walk():
            fn = p.get_filename() or ''
            if fn.endswith('.xlsx') and '日報' in fn:
                try:
                    yield day, reg, d43(p.get_content())
                except Exception:
                    pass
                break


def from_daily_dir():
    """mini：信件備份另存的 日報附件/日期/他區_南區_南區日報_日期.xlsx。"""
    for fp in sorted((ARCHIVE / '日報附件').glob('*/他區_*.xlsx')):
        m = re.match(r'他區_([\u4e00-\u9fff]{1,4}區)_.*_(\d{8})\.xlsx$', fp.name)
        if m:
            try:
                yield m.group(2), m.group(1), d43(fp.read_bytes())
            except Exception:
                pass


def from_archive():
    msgs = ARCHIVE / 'msgs'
    if not msgs.exists():
        return
    for p in msgs.glob('*.json'):
        try:
            d = json.loads(p.read_text(encoding='utf-8'))
        except Exception:
            continue
        subj = str(d.get('subject', ''))
        if OWN.search(subj) or not SUBJ.search(subj) or not d.get('saved_files'):
            continue
        day, reg = region_day(subj, str(d.get('received', '')))
        for f in d['saved_files']:
            fp = Path(f) if os.path.isabs(f) else ARCHIVE / f
            if fp.suffix == '.xlsx' and fp.exists() and day and reg:
                try:
                    yield day, reg, d43(fp.read_bytes())
                except Exception:
                    pass


def main():
    since = sys.argv[1] if len(sys.argv) > 1 else '2026-05-01'
    have, rows = load_existing()
    n = 0
    for src in (from_apple_mail(since), from_daily_dir(), from_archive()):
        for day, reg, text in src:
            if text and (day, reg) not in have:
                have.add((day, reg))
                rows.append((day, reg, text))
                n += 1
    rows.sort()
    OUT.write_text(''.join(f'## {d} {r}\n{t}\n\n' for d, r, t in rows), encoding='utf-8')
    print(f'新增 {n} 則，共 {len(rows)} 則 → {OUT}')


if __name__ == '__main__':
    main()
