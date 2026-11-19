"""EPB 匯出檔同步到 mini 後推 Discord：每份檔涵蓋哪天、幾店、能不能用，以及盲測要的日期齊了沒。

筆電 scripts/EPB匯出_同步.sh 傳完新檔後，用 ssh 在 mini 執行；新檔名從 stdin 一行一個傳進來（檔名有空白、括號）。
webhook 用 mini 自己的（筆電不放 webhook）。--dry 只印不推。
"""
import datetime as dt
import json
import os
import sys
import urllib.request
from pathlib import Path

import export_loader as X
from epb_offline import ALL, EXPORT_DIR, scan

WEBHOOK = Path('~/.config/studioa/mail_digest_discord_webhook.txt').expanduser()
BLIND_DIR = Path(os.environ.get('BLIND_OUT') or '~/工具中心/輸出/核對').expanduser() / '自動盲測'   # 同 日報盲測_自動.OUT_BASE


def md(d):
    return f'{d.month}/{d.day}'


def file_line(name):
    p = EXPORT_DIR / name
    why = X.check_file(p)
    if why:
        return f'⚠️ {name}：不能用（{why}）'
    a, b, st, t = X._scope(p)
    end = X.counted_end(b, t)
    span = md(a) if a == b else f'{md(a)}～{md(b)}'
    line = f'✅ {span}｜{len(st)} 店'
    if end < b:                                   # 匯出當天、還沒到閉店時間
        line = (f'⚠️ {span}｜{len(st)} 店：{md(b)} 是 {t[11:16]} 匯出，未過 {X.CLOSE_AFTER} 不算'
                + (f'（只算到 {md(end)}）' if end >= a else '') + f'，{X.CLOSE_AFTER} 以後或隔天再匯一次')
    return line


def message(names, now=None):
    now = now or dt.datetime.now()
    good, _ = scan()
    # 盲測要的範圍：閉店後收到就是「明天 12:30 寫今天」，否則是「今天 12:30 寫昨天」
    hm = now.strftime('%H:%M')
    target = now.date() if hm >= X.CLOSE_AFTER else now.date() - dt.timedelta(days=1)
    lack = X.missing_days(good, target.replace(day=1), target, ALL)
    lines = ['📥 **EPB 匯出檔已同步到 mini**'] + [file_line(n) for n in names]
    if target == now.date():
        run = f'明天 12:30 盲測（寫 {md(target)}）'
    elif (BLIND_DIR / target.strftime('%Y%m%d') / 'result.json').exists():   # 白天才傳到：今天那輪已經寫完
        lines.append(f'ℹ️ 今天的盲測（寫 {md(target)}）已經寫完，這份今天用不到')
        return '\n'.join(lines)
    elif hm >= '12:30':
        run = f'今天的盲測（寫 {md(target)}）正在等檔'
    else:
        run = f'今天 12:30 盲測（寫 {md(target)}）'
    if lack:
        days = sorted({d for v in lack.values() for d in v})
        lines.append(f'❗ {run}，還缺 ' + '、'.join(md(d) for d in days[:5]) + ('…' if len(days) > 5 else '')
                     + f'（{len(lack)} 店）')
    else:
        lines.append(f'👌 {run}，資料已齊：{md(target.replace(day=1))}～{md(target)} 13 店都有'
                     + ('，10 分鐘內開始寫' if '正在等檔' in run else ''))
    return '\n'.join(lines)


def main():
    names = [x.strip() for x in sys.stdin.read().splitlines() if x.strip()]
    msg = message(names)
    print(msg)
    if '--dry' not in sys.argv:
        req = urllib.request.Request(WEBHOOK.read_text().strip(), data=json.dumps({'content': msg}).encode(),
                                     headers={'Content-Type': 'application/json', 'User-Agent': 'epb-export-sync'})
        urllib.request.urlopen(req, timeout=20)


if __name__ == '__main__':
    main()
