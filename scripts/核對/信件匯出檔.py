"""同事代傳：白名單同事寄到 Bob 公司信箱的 EPB 匯出檔 → 放進 mini 的匯出資料夾，推 Discord。

2026-11-19 Bob 出國時由同權限同事匯出、寄信代傳（auto-memory epb_gui_export_source）。
mini 排程 com.bob.epb-mail-import 每 10 分鐘：先跑 mail_archive.py --imap 備份收件匣，再跑這支。
- 只看白名單寄件人（~/.config/studioa/epb_export_senders.txt，一行一個 email；檔案空的就什麼都不做）
- 只看近 3 天備份的信、只取 .xlsx 附件；不讀信件內文
- 收到後用 Bob 公司信箱回信給寄件人（SMTP 465）說明收到哪些、能不能用；同一個檔（內容雜湊）只處理一次；不能用的檔（不是匯出檔、缺欄位、有篩選、單據日期看不懂）不放進去，Discord 說原因
--dry 只印不推、不搬檔。
"""
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import smtplib
import ssl
import sys
import time
import urllib.request
from email.message import EmailMessage
from pathlib import Path

import export_loader as X
from epb_offline import EXPORT_DIR

ARCHIVE = Path(os.environ.get('MAIL_ARCHIVE_DIR') or Path(__file__).resolve().parent.parent / '信件追蹤' / 'archive')
SENDERS = Path(os.environ.get('EPB_EXPORT_SENDERS') or '~/.config/studioa/epb_export_senders.txt').expanduser()
STATE = EXPORT_DIR / '_mail_imported.json'
DAYS = 3
IMAP_ENV = Path('~/.config/studioa/imap.env').expanduser()     # 公司信箱帳號（mail_archive 同一組）；回信走 SMTP 465


def senders():
    if not SENDERS.exists():
        return set()
    return {l.split('#')[0].strip().lower() for l in SENDERS.read_text().splitlines() if l.split('#')[0].strip()}


def recent_msgs(days=DAYS):
    """近 days 天備份的信（用 json 檔修改時間篩，免得每次讀一萬多封）。"""
    cut = time.time() - days * 86400
    with os.scandir(ARCHIVE / 'msgs') as it:
        for e in it:
            if e.name.endswith('.json') and e.stat().st_mtime >= cut:
                try:
                    yield json.loads(Path(e.path).read_text(encoding='utf-8'))
                except Exception:
                    continue


def run(dry=False):
    allow = senders()
    if not allow:
        return [], [], []
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    good, bad, replies, ignored = [], [], [], False
    for rec in sorted(recent_msgs(), key=lambda r: r.get('received', '')):
        who = (rec.get('sender') or '').lower()
        if who not in allow:
            continue
        items = []                                         # 這封信新處理的檔：(原檔名, 說明, 能不能用)
        for fn in rec.get('saved_files') or []:
            src = ARCHIVE / '附件' / rec['key'] / fn
            if not fn.lower().endswith('.xlsx') or not src.exists():
                continue
            h = hashlib.sha1(src.read_bytes()).hexdigest()
            if h in state:
                continue
            tmp = EXPORT_DIR / f'.mail_{h[:12]}.xlsx'       # 點開頭：掃描與同步都會跳過
            shutil.copyfile(src, tmp)
            why = X.check_file(tmp)
            if why and why.startswith(('不是', '讀不到')) and not fn.startswith('銷售分析作業'):
                # 同事寄的其他 xlsx（例：Yuni 每天寄的 Daily report）不是代傳檔：不通知、不回信，記下來以後不再看
                tmp.unlink()
                if not dry:
                    state[h] = {'file': None, 'sender': who, 'received': rec.get('received'), 'why': '略過（不是匯出檔）',
                                'at': dt.datetime.now().strftime('%F %T')}
                    ignored = True
                continue
            stamp = rec.get('received', '')[5:16].replace('-', '').replace(' ', '_').replace(':', '')
            name = f'信件_{stamp}_{who.split("@")[0]}_{re.sub(r"[/:]", "_", fn)}'
            items.append((fn, why or describe(tmp), not why))
            if why or dry:
                tmp.unlink()
            else:
                tmp.replace(EXPORT_DIR / name)
            if why:
                bad.append((fn, who, why))
            else:
                good.append((name, who))
            if not dry:
                state[h] = {'file': None if why else name, 'sender': who, 'received': rec.get('received'),
                            'why': why, 'at': dt.datetime.now().strftime('%F %T')}
        if items:
            replies.append((rec, items))
    if not dry and (good or bad or ignored):
        STATE.write_text(json.dumps(state, ensure_ascii=False, indent=1))
    return good, bad, replies


def describe(path):
    """可用的匯出檔 → 「11/19（13 店）」；當天未過 22:30 匯出的另外說明不採用。"""
    a, b, st, t = X._scope(path)
    md = lambda d: f'{d.month}/{d.day}'
    span = md(a) if a == b else f'{md(a)}～{md(b)}'
    end = X.counted_end(b, t)
    if end < b:
        return (f'{span}（{len(st)} 店），但 {md(b)} 是 {t[11:16]} 匯出、未過 {X.CLOSE_AFTER}，這天的資料不會採用，'
                f'請 {X.CLOSE_AFTER} 以後或隔天再匯一次')
    return f'{span}（{len(st)} 店），資料可以使用'


def reply_body(items):
    lines = ['你好，', '', '已收到你寄來的 EPB 匯出檔：']
    for fn, info, ok in items:
        if ok:
            lines.append(f'- {fn}：{info}')
        else:
            lines.append(f'- {fn}：無法使用（{info}）。請確認條件只設「單據日期 BETWEEN 起訖同一天」和 13 家店、'
                         '不加其他篩選，重新匯出後再寄一次')
    return '\n'.join(lines + ['', '謝謝幫忙！', '', '（此信由系統自動回覆）'])


def send_reply(rec, body):
    """用 Bob 的公司信箱回信給寄件人（接在原信串）。"""
    env = dict(l.strip().split('=', 1) for l in IMAP_ENV.read_text().splitlines() if '=' in l)
    env = {k: v.strip('"\'') for k, v in env.items()}
    m = EmailMessage()
    m['From'], m['To'] = env['work_USER'], rec['sender']
    subj = (rec.get('subject') or '').strip()
    m['Subject'] = subj if subj.lower().startswith('re:') else f'Re: {subj or "EPB 匯出檔"}'
    if rec.get('message_id'):
        m['In-Reply-To'] = m['References'] = rec['message_id']
    m.set_content(body)
    ctx = ssl.create_default_context()
    ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE   # 公司自簽憑證（同 mail_archive）
    try:
        ctx.set_ciphers('DEFAULT@SECLEVEL=1')                  # Homebrew OpenSSL 要降等級；系統 LibreSSL 不支援也不需要
    except ssl.SSLError:
        pass
    with smtplib.SMTP_SSL(env['work_HOST'], 465, timeout=30, context=ctx) as s:
        s.login(env['work_USER'], env['work_PASSWORD'])
        s.send_message(m)


def message(good, bad):
    import importlib
    N = importlib.import_module('匯出檔通知')
    who = '、'.join(sorted({w.split('@')[0] for _, w in good} | {w.split('@')[0] for _, w, _ in bad}))
    lines = [f'📧 **同事代傳的 EPB 匯出檔**（{who}）']
    if good:
        lines += N.message([n for n, _ in good]).splitlines()[1:]     # 去掉「已同步到 mini」標題，其餘照同步通知
    for fn, _, why in bad:
        lines.append(f'⚠️ {fn}：不能用（{why}），請重新匯出再寄')
    return '\n'.join(lines)


def main():
    dry = '--dry' in sys.argv
    good, bad, replies = run(dry)
    if not good and not bad:
        return
    msg = message(good, bad)
    for rec, items in replies:                             # 每封信回一次，說明收到哪些、能不能用
        body = reply_body(items)
        if dry:
            print(f'（--dry 不寄）回信給 {rec["sender"]}：\n{body}')
            continue
        try:
            send_reply(rec, body)
            msg += f'\n✉️ 已回信給 {rec["sender"]}'
        except Exception as e:
            msg += f'\n⚠️ 回信給 {rec["sender"]} 失敗（{type(e).__name__}），請自己回覆'
    print(dt.datetime.now().strftime('%F %T'), msg.replace('\n', ' ｜ '))
    if not dry:
        url = Path('~/.config/studioa/mail_digest_discord_webhook.txt').expanduser().read_text().strip()
        req = urllib.request.Request(url, data=json.dumps({'content': msg}).encode(),
                                     headers={'Content-Type': 'application/json', 'User-Agent': 'epb-mail-import'})
        urllib.request.urlopen(req, timeout=20)


if __name__ == '__main__':
    main()
