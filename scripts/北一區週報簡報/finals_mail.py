"""從信箱取回報告人寄出的週報定稿（北二區：Yuni 每週寄「北二區 27Q1W2 週報」附 .key）。

存成跟本工具範本相同的檔名（例：北二區週報 27Q1W02.key）放在 <輸出>/定稿信件/：
build_deck 拿它當下週範本（保留報告人手改的內容），並把裡面的評語當定稿來學寫法。
同一週寄多封（「以此封為主」）取最後收到的那封。

    python3 finals_mail.py --region 北二區     # 手動補抓全部歷史定稿
"""
import email
import glob
import re
from email import policy
from pathlib import Path

import build_deck as b


def _emlx(rowid):
    # Mail 依 ROWID 千位數倒序分層：194192 → Data/4/9/1/Messages/194192.emlx
    sub = '/'.join(reversed(str(rowid // 1000)))
    hits = glob.glob(str(b.MAIL_BASE / '*' / '*.mbox' / '*' / 'Data' / sub / 'Messages' / f'{rowid}.emlx'))
    return Path(hits[0]) if hits else None


def _key_attachment(rowid):
    p = _emlx(rowid)
    if not p:
        return None
    raw = p.read_bytes()
    m = email.message_from_bytes(raw[raw.index(b'\n') + 1:], policy=policy.default)
    for part in m.walk():
        if (part.get_filename() or '').lower().endswith('.key'):
            data = part.get_payload(decode=True)
            if data and data[:2] == b'PK':
                return data
    for f in glob.glob(str(p.parent.parent / 'Attachments' / str(rowid) / '**' / '*.key'), recursive=True):
        return Path(f).read_bytes()
    return None


def fetch_finals():
    """抓本區所有定稿信（已存過的不重抓）。回傳 {週次: 路徑}，例 {'27Q1W02': Path(...)}。"""
    cfg = b.R['final_mail']
    out_dir = b.R['final_dir']
    out_dir.mkdir(parents=True, exist_ok=True)
    con = b._mail_db()
    rows = con.execute(
        "select m.ROWID, s.subject from messages m join subjects s on m.subject=s.ROWID "
        "join addresses a on m.sender=a.ROWID where a.address=? order by m.date_received",
        (cfg['sender'],)).fetchall()
    latest = {}
    for rid, subj in rows:
        mt = re.search(cfg['subject'], subj or '')
        if mt:
            latest[f'{mt[1]}Q{mt[2]}W{int(mt[3]):02d}'] = rid       # 依收信時間排序，後面的蓋掉前面的
    got = {}
    for label, rid in latest.items():
        f = out_dir / f"{b.R['deck_prefix']} {label}.key"
        if not f.exists():
            data = _key_attachment(rid)
            if not data:
                continue
            f.write_bytes(data)
            b.log(f'定稿信件：取回 {f.name}')
        got[label] = f
    return got


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('--region', default='北二區')
    b.configure(ap.parse_args().region)
    for k, v in sorted(fetch_finals().items()):
        print(k, v)
