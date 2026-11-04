#!/usr/bin/env python3
"""信件備份：把公司信存成本機 JSON，給 mail_digest 和之後的門市知識庫讀。

公司 IMAP 只留最近約 24 小時的信（容量滿就刪），所以要勤抓、自己留底。
一封信一個檔：archive/msgs/<key>.json（key＝Message-ID 雜湊，兩個來源同一封信不會重複）。
附件只留文件類（PDF／Excel／Word／PPT／CSV，≤15MB），圖片不留，存在 archive/附件/<key>/。

用法：
  python3 mail_archive.py --imap          抓 IMAP 收件匣（排程每小時跑；要 Homebrew python，系統 python 的 LibreSSL 連不上）
  python3 mail_archive.py --apple 200     從這台 MacBook 的 Apple Mail 匯入近 200 天（補舊信／補缺口；要在有「完整磁碟取用」的終端機跑）
"""
import argparse
import email
import email.policy
import hashlib
import html
import imaplib
import json
import os
import re
import shutil
import sqlite3
import ssl
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta
from email.utils import parseaddr
from pathlib import Path

HERE = Path(__file__).resolve().parent
ARCHIVE = HERE / 'archive'
MSGS = ARCHIVE / 'msgs'
ATT = ARCHIVE / '附件'
STATE = ARCHIVE / 'state.json'
IMAP_ENV = Path.home() / '.config/studioa/imap.env'
MAIL_BASE = Path.home() / 'Library/Mail/V10'
KEEP_EXT = ('.pdf', '.xlsx', '.xls', '.docx', '.doc', '.pptx', '.csv')
MAX_ATT = 15 * 1024 * 1024
# 日盤表、檢查表、週報這類例行報表佔掉 2/3 空間又不是 SOP → 不留檔（檔名仍記在 files）
REPORT_NAME = re.compile(r'盤|檢查表|週報|日報|寄銷|報表|業績|Daily|Weekly', re.I)


def keep_attachment(fn, sender):
    from mail_digest import STORE_BOXES   # 門市信箱寄的附件都是例行回報
    return (fn.lower().endswith(KEEP_EXT) and not REPORT_NAME.search(fn)
            and sender.split('@')[0].lower() not in STORE_BOXES)


# ---------- 解析 ----------

def _html_to_text(s):
    s = re.sub(r'(?is)<(script|style).*?</\1>', '', s)
    s = re.sub(r'(?i)<br\s*/?>|</p>|</div>|</tr>|</li>', '\n', s)
    s = re.sub(r'(?s)<[^>]+>', '', s)
    return html.unescape(s)


def msg_key(msg):
    # Outlook 有時不同封信共用同一個 Message-ID（<!&!…@studioa.com.tw>），所以加上 Date/From/Subject
    k = '|'.join(str(msg.get(h, '')).strip() for h in ('Message-ID', 'Date', 'From', 'Subject'))
    return hashlib.sha1(k.encode('utf-8', 'replace')).hexdigest()[:20]


def save(raw, received, src, sender=None, sender_name=None, subject=None):
    """解析一封原始信，存 JSON＋文件附件。已存在就略過，回傳 True 表示新存。"""
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    key = msg_key(msg)
    out = MSGS / f'{key}.json'
    if out.exists():
        return False
    name, addr = parseaddr(str(msg.get('From', '')))
    sender = (sender or addr).lower()
    plain, htm, files, kept = None, None, [], []
    for part in msg.walk():
        fn = part.get_filename()
        if fn:
            files.append(fn)
            if keep_attachment(fn, sender):
                try:
                    data = part.get_payload(decode=True) or b''
                except Exception:
                    data = b''
                if 0 < len(data) <= MAX_ATT:
                    d = ATT / key
                    d.mkdir(parents=True, exist_ok=True)
                    safe = re.sub(r'[/\\:\x00]', '_', fn).strip().lstrip('.') or 'attachment'
                    (d / safe).write_bytes(data)
                    kept.append(safe)
            continue
        ctype = part.get_content_type()
        try:
            if ctype == 'text/plain' and plain is None:
                plain = part.get_content()
            elif ctype == 'text/html' and htm is None:
                htm = part.get_content()
        except Exception:
            continue
    text = plain if plain and plain.strip() else _html_to_text(htm or '')
    text = re.sub(r'[ \t　\xa0]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n', text).strip()
    rec = {
        'key': key, 'message_id': (msg.get('Message-ID') or '').strip(),
        'received': received.strftime('%Y-%m-%d %H:%M:%S'),
        'sender': sender, 'sender_name': sender_name if sender_name is not None else name,
        'subject': subject if subject is not None else str(msg.get('Subject', '')).strip(),
        'body': text, 'files': files, 'saved_files': kept, 'src': src,
    }
    MSGS.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix('.tmp')
    tmp.write_text(json.dumps(rec, ensure_ascii=False))
    tmp.replace(out)
    return True


def _state():
    return json.loads(STATE.read_text()) if STATE.exists() else {}


def _write_state(s):
    ARCHIVE.mkdir(parents=True, exist_ok=True)
    STATE.write_text(json.dumps(s, ensure_ascii=False, indent=1))


def _notify(text):
    subprocess.run(['osascript', '-e', f'display notification "{text}" with title "信件備份"'],
                   capture_output=True, timeout=30)


# ---------- IMAP ----------

def from_imap():
    env = dict(l.strip().split('=', 1) for l in IMAP_ENV.read_text().splitlines() if '=' in l)
    env = {k: v.strip('"\'') for k, v in env.items()}
    # 公司 Dovecot 自簽憑證＋弱 DH：要降安全等級才連得上（同 FLUX vault email_ops.py）
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    ctx.set_ciphers('DEFAULT@SECLEVEL=1')
    M = imaplib.IMAP4(env['work_HOST'], int(env['work_PORT']), timeout=60)
    M.starttls(ssl_context=ctx)
    M.login(env['work_USER'], env['work_PASSWORD'])
    n = int(M.select('INBOX', readonly=True)[1][0])
    new = seen = failed = 0
    oldest = None
    if n:
        typ, data = M.fetch(f'1:{n}', '(UID INTERNALDATE BODY.PEEK[HEADER.FIELDS (MESSAGE-ID DATE FROM SUBJECT)])')
        heads = [x for x in data if isinstance(x, tuple)]
        for meta, hdr in heads:
            meta = meta.decode()
            uid = re.search(r'UID (\d+)', meta).group(1)
            t = imaplib.Internaldate2tuple(meta.encode())
            received = datetime(*t[:6])
            oldest = min(oldest or received, received)
            key = msg_key(email.message_from_bytes(hdr, policy=email.policy.default))
            if (MSGS / f'{key}.json').exists():
                seen += 1
                continue
            # 單封解析失敗不能中斷整輪：那封沒存，下一小時又重試又中斷，備份就永遠卡住
            try:
                typ, full = M.uid('FETCH', uid, '(BODY.PEEK[])')
                raw = next(x[1] for x in full if isinstance(x, tuple))
                new += save(raw, received, 'imap')
            except Exception as e:
                failed += 1
                print('  存信失敗', uid, type(e).__name__, e, file=sys.stderr)
    M.logout()

    st = _state()
    last = st.get('last_imap')
    # 伺服器上最舊那封比上次成功抓的時間還晚 → 中間可能有信已被伺服器刪掉
    if last and oldest and oldest > datetime.fromisoformat(last):
        gap = f'{last[:16]} ～ {oldest:%Y-%m-%d %H:%M}'
        st['gap'] = gap
        print('⚠️ 可能有缺口：', gap, '→ 在 MacBook 終端機跑 mail_archive.py --apple 7，'
              '再 rsync -a --ignore-existing archive/msgs/ archive/附件/ 到 mini 同一路徑')
        _notify(f'可能漏信 {gap}，請在 MacBook 跑 --apple 補')
    st['last_imap'] = datetime.now().isoformat(timespec='seconds')
    _write_state(st)
    print(f'IMAP 收件匣 {n} 封：新存 {new}、已有 {seen}' + (f'、失敗 {failed}' if failed else ''))


# ---------- Apple Mail（只在這台 MacBook） ----------

def from_apple(days):
    tmp = tempfile.mkdtemp()
    try:
        for suffix in ('', '-wal', '-shm'):
            src = MAIL_BASE / 'MailData' / ('Envelope Index' + suffix)
            if src.exists():
                shutil.copy2(src, os.path.join(tmp, 'ei' + suffix))
        con = sqlite3.connect(os.path.join(tmp, 'ei'))
        lo = (datetime.now() - timedelta(days=days)).timestamp()
        rows = con.execute(
            """SELECT m.ROWID, m.date_received, a.address, COALESCE(a.comment,''), s.subject,
                      COALESCE(m.subject_prefix,''), mb.url
               FROM messages m
               JOIN addresses a ON a.ROWID = m.sender
               JOIN subjects s ON s.ROWID = m.subject
               JOIN mailboxes mb ON mb.ROWID = m.mailbox
               WHERE m.deleted = 0 AND m.date_received >= ?""", (lo,)).fetchall()
        con.close()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    paths = {}
    for f in MAIL_BASE.glob('*/**/Messages/*.emlx'):
        rid = f.name.split('.')[0]
        if rid not in paths or not f.name.endswith('.partial.emlx'):
            paths[rid] = f
    new = skip = miss = 0
    for rid, ts, addr, name, subj, prefix, url in rows:
        low = url.lower()
        if any(k in low for k in ('sent', 'draft', 'trash', 'junk', '寄件', '草稿', '垃圾')):
            continue
        p = paths.get(str(rid))
        if not p:
            miss += 1
            continue
        raw = p.read_bytes()
        nl = raw.find(b'\n')
        try:
            raw = raw[nl + 1: nl + 1 + int(raw[:nl])]
        except ValueError:
            pass
        try:
            ok = save(raw, datetime.fromtimestamp(ts), 'apple', addr, name, (prefix + subj).strip())
        except Exception as e:
            print('  讀信失敗', rid, e, file=sys.stderr)
            continue
        new += ok
        skip += not ok
    print(f'Apple Mail 近 {days} 天：新存 {new}、已有 {skip}、找不到檔 {miss}')
    st = _state()
    st.pop('gap', None)
    _write_state(st)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--imap', action='store_true')
    ap.add_argument('--apple', type=int, metavar='DAYS')
    a = ap.parse_args()
    if a.apple:
        from_apple(a.apple)
    if a.imap:
        from_imap()
    if not (a.apple or a.imap):
        ap.print_help()


if __name__ == '__main__':
    main()
