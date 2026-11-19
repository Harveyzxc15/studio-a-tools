"""同事代傳：白名單同事寄到 Bob 公司信箱的 EPB 匯出檔 → 放進 mini 的匯出資料夾，推 Discord。

2026-11-19 Bob 出國時由同權限同事匯出、寄信代傳（auto-memory epb_gui_export_source）。
mini 排程 com.bob.epb-mail-import 每 10 分鐘：先跑 mail_archive.py --imap 備份收件匣，再跑這支。
- 只看白名單寄件人（~/.config/studioa/epb_export_senders.txt，一行一個 email；檔案空的就什麼都不做）
- 只看近 3 天備份的信、只取 .xlsx 附件；不讀信件內文
- 同一個檔（內容雜湊）只處理一次；不能用的檔（不是匯出檔、缺欄位、有篩選、單據日期看不懂）不放進去，Discord 說原因
--dry 只印不推、不搬檔。
"""
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import sys
import time
import urllib.request
from pathlib import Path

import export_loader as X
from epb_offline import EXPORT_DIR

ARCHIVE = Path(os.environ.get('MAIL_ARCHIVE_DIR') or Path(__file__).resolve().parent.parent / '信件追蹤' / 'archive')
SENDERS = Path(os.environ.get('EPB_EXPORT_SENDERS') or '~/.config/studioa/epb_export_senders.txt').expanduser()
STATE = EXPORT_DIR / '_mail_imported.json'
DAYS = 3


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
        return [], []
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    good, bad = [], []
    for rec in sorted(recent_msgs(), key=lambda r: r.get('received', '')):
        who = (rec.get('sender') or '').lower()
        if who not in allow:
            continue
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
            stamp = rec.get('received', '')[5:16].replace('-', '').replace(' ', '_').replace(':', '')
            name = f'信件_{stamp}_{who.split("@")[0]}_{re.sub(r"[/:]", "_", fn)}'
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
    if not dry and (good or bad):
        STATE.write_text(json.dumps(state, ensure_ascii=False, indent=1))
    return good, bad


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
    good, bad = run(dry)
    if not good and not bad:
        return
    msg = message(good, bad)
    print(dt.datetime.now().strftime('%F %T'), msg.replace('\n', ' ｜ '))
    if not dry:
        url = Path('~/.config/studioa/mail_digest_discord_webhook.txt').expanduser().read_text().strip()
        req = urllib.request.Request(url, data=json.dumps({'content': msg}).encode(),
                                     headers={'Content-Type': 'application/json', 'User-Agent': 'epb-mail-import'})
        urllib.request.urlopen(req, timeout=20)


if __name__ == '__main__':
    main()
