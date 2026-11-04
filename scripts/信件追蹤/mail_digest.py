#!/usr/bin/env python3
"""信件追蹤：從信件備份（mail_archive.py）抽出交辦截止日、SOP 最新版、活動期間。

流程：讀 archive/msgs → 過濾雜訊 → 規則抓主旨日期 → 抓不到的總部信才送 Claude（claude -p）
      → 同主題歸成系列 → 產 snapshot.json →（--push）推 sa-north-hub、（--calendar）寫行事曆 SA。

用法：
  python3 mail_digest.py                 只產 snapshot.json（預設近 45 天）
  python3 mail_digest.py --push --calendar
  python3 mail_digest.py --days 60 --no-claude
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
ARCHIVE = HERE / 'archive/msgs'              # mail_archive.py 存的信
CACHE = HERE / 'claude_cache.json'          # 信件 key → Claude 判讀結果，每封只問一次
CAL_STATE = HERE / 'calendar_state.json'     # 已寫進行事曆的截止日，避免重複建
SNAPSHOT = HERE / 'snapshot.json'
TOKEN_FILE = Path.home() / '.config/studioa/epb_push_token'
CLAUDE_TOKEN = Path.home() / '.config/studioa/claude_oauth_token'
DISCORD_HOOK = Path.home() / '.config/studioa/mail_digest_discord_webhook.txt'
HUB = 'https://sa-north-hub.fly.dev'
CALENDAR = 'SA'

# ---------- 雜訊過濾 ----------
NOISE_SENDERS = {
    'dss@studioa.com.tw', 'jobbank@104.com.tw', 'guestbook_support@group.apple.com',
    'noreply@studioack.itpison.com', 'notice@studioa.com.tw', 'service@lt365.com.tw',
    'jnoreply@studioa.com.tw',
}
# 系統催辦、人事個資（請假／離職／考核／面試）不上雲端
NOISE_SUBJECT = re.compile(
    r'daily ?report|日報|銷報|週報\(滿意度\)|Callcard|新聞摘錄|待簽|加班單|請假|離職|考核|面試|OTP|驗證信|'
    r'垃圾郵件|未結案通知|訂單待確認|Undeliverable|自動回覆|Automatic reply|保險批改', re.I)
# 門市信箱（寄件者是門市）→ 歸「門市來信」，不送 Claude
STORE_BOXES = {
    'shihlin': '士林', 'breeze': '微風', 'miramar': '美麗華', 'apollo': '阿波羅',
    'dayeh': '大葉高島屋', 'luodong': '羅東', 'yonghe': '永和', 'banqiaoeslite': '板橋誠品',
    'ximen': '新西門', 'hualien': '花蓮', 'banqiaofe21': '板橋遠百', 'honhui': '新莊宏匯',
    'yulongcity': '新店裕隆城',
}

SOP_WORDS = re.compile(r'SOP|規範|流程|須知|作業|辦法|規則|說明|格式|表單調整|操作|手冊|注意事項')
EVENT_WORDS = re.compile(r'活動|出清|特賣|Roadshow|回購|優惠|促銷|折扣|贈品|加購|檔期|方案')
TASK_WORDS = re.compile(r'請.{0,12}(回覆|填寫|寄出|確認|完成|繳交|提供|補上|處理)|回覆|填寫|寄回|前寄出|截止')

# ---------- 日期規則 ----------
WEEKDAY = '一二三四五六日'
RE_DEADLINE = [
    # 10/21前、請於10/15前回覆、10/19 前 準時寄出、(10/21)前
    re.compile(r'(?<![\d/])(\d{1,2})\s*/\s*(\d{1,2})\s*(?:[（(][一二三四五六日][)）])?\s*(?:\d{1,2}[:：]\d{2}|\d{1,2}點)?\s*(?:以?前|截止)'),
    re.compile(r'(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*(?:[（(][一二三四五六日][)）])?\s*(?:以?前|截止)'),
    re.compile(r'截止[日期時間]*\s*[:：]?\s*(?:\d{4}\s*[/.-]\s*)?(\d{1,2})\s*[/.-]\s*(\d{1,2})'),
]
RE_WEEKDAY_DEADLINE = re.compile(r'(?:本)?週([一二三四五六日])\s*(?:\d{1,2}\s*點|下午|中午|上午)?\s*(?:\d{1,2}\s*點)?\s*以?前')
RE_PERIOD = re.compile(
    r'(?:(\d{4})\s*[/.-]\s*)?(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*(?:[（(][一二三四五六日][)）])?\s*[~～\-–至到]\s*'
    r'(?:(\d{4})\s*[/.-]\s*)?(\d{1,2})\s*[/.-]\s*(\d{1,2})')


def _mkdate(m, d, ref):
    """月日 → 日期，年份取最接近收信日的那年（12 月收信寫 1/5 → 隔年）。"""
    try:
        cands = [date(ref.year + k, int(m), int(d)) for k in (-1, 0, 1)]
    except ValueError:
        return None
    return min(cands, key=lambda x: abs((x - ref).days))


def find_deadline(text, ref):
    for rx in RE_DEADLINE:
        mm = rx.search(text)
        if mm:
            dl = _mkdate(mm.group(1), mm.group(2), ref)
            if dl and -3 <= (dl - ref).days <= 120:
                return dl
    mm = RE_WEEKDAY_DEADLINE.search(text)
    if mm:
        wd = WEEKDAY.index(mm.group(1))
        return ref + timedelta(days=(wd - ref.weekday()) % 7)
    return None


def find_period(text, ref):
    mm = RE_PERIOD.search(text)
    if not mm:
        return None
    a = _mkdate(mm.group(2), mm.group(3), ref)
    b = _mkdate(mm.group(5), mm.group(6), ref)
    if not a or not b:
        return None
    if b < a:
        b = date(b.year + 1, b.month, b.day)
    if (b - a).days > 400:
        return None
    return a, b


# ---------- 讀信（從 mail_archive.py 的本機備份） ----------

def load_messages(days):
    lo = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d %H:%M:%S')
    out = []
    for f in ARCHIVE.glob('*.json'):
        r = json.loads(f.read_text())
        if r['received'] < lo:
            continue
        # 砍掉引用的舊信（回信串）
        body = re.split(r'\n\s*(?:-{2,}\s*(?:原始郵件|Original Message)|寄件者[:：]|From:\s|在 .{5,60}寫道)', r['body'])[0]
        out.append({'id': r['key'], 'received': datetime.strptime(r['received'], '%Y-%m-%d %H:%M:%S'),
                    'sender': r['sender'], 'sender_name': r['sender_name'], 'subject': r['subject'],
                    'body': body.strip(), 'files': r['files']})
    out.sort(key=lambda m: m['received'])
    return out


# ---------- 分類與系列 ----------

def classify(subject, body, sender):
    local = sender.split('@')[0].lower()
    if local in STORE_BOXES:
        return 'store'
    head = subject + '\n' + body[:400]
    if SOP_WORDS.search(subject):
        return 'sop'
    if EVENT_WORDS.search(subject):
        return 'event'
    if TASK_WORDS.search(head):
        return 'task'
    return 'notice'


def topic_key(subject):
    s = re.sub(r'^\s*((RE|FW|FWD|回覆|轉寄)\s*[:：]\s*)+', '', subject, flags=re.I)
    s = re.sub(r'[［\[【(（]?\s*(更新|修正|補充|重寄|再次提醒|提醒)\s*[］\]】)）]?', '', s)
    s = re.sub(r'\d{2,4}\s*[/.-]\s*\d{1,2}\s*[/.-]\s*\d{1,2}|\d{8}|\d{1,2}/\d{1,2}(\s*[~～-]\s*\d{1,2}/\d{1,2})?|\d{3,4}(?=\s)', '', s)
    s = re.sub(r'\s*[-_－]\s*(APR|AAR|CEC|A ?Store|士林|微風|美麗華|阿波羅|高島屋|羅東|永和|誠品|西門|花蓮|遠百|宏匯|裕隆城)\s*$', '', s)
    s = re.sub(r'[\s\-－_~～—–:：,，。.！!🔺]+', '', s)
    return s[:40]


# ---------- Claude 補判 ----------
PROMPT = """你是門市區經理的助理。以下是公司內部信件（JSON 陣列），每封判讀：
- summary：一句話（30 字內）說這封信要門市／區經理做什麼或知道什麼
- deadline：門市要完成動作的截止日 YYYY-MM-DD，沒有就 null（只算「要做事的期限」，活動結束日不算）
- period_start、period_end：活動或規範的生效期間 YYYY-MM-DD，沒有就 null
- action：要做的事（20 字內），純公告不用做事就 null
- changes：若是 SOP／規範更新，說改了什麼（40 字內），否則 null
年份依收信日推算。只輸出 JSON 陣列，每個元素 {"id":..., "summary":..., "deadline":..., "period_start":..., "period_end":..., "action":..., "changes":...}，不要其他文字。

信件：
"""


def ask_claude(batch):
    payload = [{'id': m['id'], 'received': m['received'].strftime('%Y-%m-%d'), 'subject': m['subject'],
                'body': m['body'][:2500], 'attachments': m['files']} for m in batch]
    try:
        # 走 claude 登入的訂閱額度，不用 ANTHROPIC_API_KEY（那把 key 沒額度）；launchd 的 PATH 沒有 ~/.local/bin
        env = {k: v for k, v in os.environ.items() if k != 'ANTHROPIC_API_KEY'}
        if CLAUDE_TOKEN.exists():   # mini：claude setup-token 產的長效憑證，背景排程用
            env['CLAUDE_CODE_OAUTH_TOKEN'] = CLAUDE_TOKEN.read_text().strip()
        r = subprocess.run([shutil.which('claude') or str(Path.home() / '.local/bin/claude'),
                            '-p', '--model', 'haiku', '--output-format', 'text'],
                           input=PROMPT + json.dumps(payload, ensure_ascii=False),
                           capture_output=True, text=True, timeout=300, env=env)
    except (OSError, subprocess.TimeoutExpired) as e:
        print('  claude 失敗：', e, file=sys.stderr)
        return {}
    txt = r.stdout.strip() or r.stderr.strip()
    mm = re.search(r'\[.*\]', txt, re.S)
    if not mm:
        print('  claude 回傳看不懂：', txt[:200], file=sys.stderr)
        return {}
    try:
        return {str(x['id']): x for x in json.loads(mm.group(0)) if isinstance(x, dict) and 'id' in x}
    except (json.JSONDecodeError, KeyError):
        return {}


def _d(s):
    try:
        return date.fromisoformat(s) if s else None
    except (TypeError, ValueError):
        return None


# ---------- 行事曆 ----------
CAL_JXA = r'''
function run(argv) {
  const [calName, title, ymd, notes] = argv;
  const app = Application('Calendar');
  const cal = app.calendars.whose({name: calName})[0];
  const [y, m, d] = ymd.split('-').map(Number);
  const start = new Date(y, m - 1, d), end = new Date(y, m - 1, d + 1);
  const ev = app.Event({summary: title, startDate: start, endDate: end, alldayEvent: true, description: notes});
  cal.events.push(ev);
  // 全天事件的預設提醒通常已是前一天 09:00；沒有才補，免得重複
  if (ev.displayAlarms().length === 0) ev.displayAlarms.push(app.DisplayAlarm({triggerInterval: -15 * 60}));
  return ev.uid();
}
'''


def sync_calendar(items):
    state = json.loads(CAL_STATE.read_text()) if CAL_STATE.exists() else {}
    today = date.today()
    added = 0
    for it in items:
        dl = it.get('deadline')
        if not dl or it['kind'] == 'store' or date.fromisoformat(dl) < today:
            continue
        key = f"{it['topic']}|{dl}"
        if key in state:
            continue
        subject = re.sub(r'^\s*((RE|FW|FWD)\s*[:：]\s*)+', '', it['subject'], flags=re.I)
        title = '⏰ ' + (it.get('action') or subject)[:60]
        notes = f"{it['subject']}\n寄件：{it['sender_name'] or it['sender']}　收信：{it['received'][:10]}\n{it.get('summary') or ''}"
        try:
            r = subprocess.run(['osascript', '-l', 'JavaScript', '-e', CAL_JXA, CALENDAR, title, dl, notes],
                               capture_output=True, text=True, timeout=60)
        except subprocess.TimeoutExpired:
            # 行事曆 App 卡住（多半是畫面上有視窗等人按）→ 這輪剩下的先跳過，下次排程再補
            print('  行事曆寫入超時，這輪先跳過剩下的', file=sys.stderr)
            break
        if r.returncode == 0:
            state[key] = r.stdout.strip()
            added += 1
        else:
            print('  行事曆寫入失敗：', r.stderr.strip()[:200], file=sys.stderr)
    CAL_STATE.write_text(json.dumps(state, ensure_ascii=False, indent=1))
    return added


# ---------- 主流程 ----------

def build(days, use_claude):
    keep = []
    for m in load_messages(days):
        if m['sender'].lower() in NOISE_SENDERS or NOISE_SUBJECT.search(m['subject']):
            continue
        if m['sender'].lower().startswith('bob.yu@'):
            continue
        ref = m['received'].date()
        m['kind'] = classify(m['subject'], m['body'], m['sender'])
        m['topic'] = topic_key(m['subject'])
        dl = find_deadline(m['subject'], ref)
        m['deadline_src'] = 'subject' if dl else None
        if not dl:
            dl = find_deadline(m['body'][:1500], ref)
            m['deadline_src'] = 'body' if dl else None
        m['deadline'] = dl
        per = find_period(m['subject'], ref)
        m['period'] = per
        keep.append(m)

    # 規則抓不到截止日的總部信 → Claude 補判（有快取）
    cache = json.loads(CACHE.read_text()) if CACHE.exists() else {}
    need = [m for m in keep if m['kind'] in ('task', 'sop', 'event') and not m['deadline'] and str(m['id']) not in cache]
    if use_claude and need:
        print(f'送 Claude 判讀 {len(need)} 封…')
        for i in range(0, len(need), 12):
            batch = need[i:i + 12]
            got = ask_claude(batch)
            if got:   # 這批有正常回來，但漏判的信記成空結果，免得每次排程都重送、一直耗額度
                for m in batch:
                    got.setdefault(str(m['id']), {'id': m['id']})
            cache.update(got)
            CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=1))
    for m in keep:
        c = cache.get(str(m['id']))
        if c:
            m['summary'], m['action'], m['changes'] = c.get('summary'), c.get('action'), c.get('changes')
            if not m['deadline'] and _d(c.get('deadline')):
                m['deadline'], m['deadline_src'] = _d(c['deadline']), 'claude'
            ps, pe = _d(c.get('period_start')), _d(c.get('period_end'))
            if not m['period'] and pe:
                m['period'] = (ps or pe, pe)
        else:
            m.setdefault('summary', None)
            m.setdefault('action', None)
            m.setdefault('changes', None)
        if not m['period']:
            m['period'] = find_period(m['body'][:800], m['received'].date()) if m['kind'] == 'event' else None

    # 同主題歸系列：最新那封當代表，記版本數
    series = {}
    for m in sorted(keep, key=lambda x: x['received']):
        series.setdefault(m['topic'], []).append(m)
    items = []
    for topic, ms in series.items():
        last = ms[-1]
        hq = [x for x in ms if x['kind'] != 'store']
        kind = hq[0]['kind'] if hq else 'store'
        dl = max((x['deadline'] for x in ms if x['deadline']), default=None)
        items.append({
            'id': last['id'], 'topic': topic, 'kind': kind, 'subject': last['subject'],
            'sender': last['sender'], 'sender_name': last['sender_name'],
            'store': STORE_BOXES.get(last['sender'].split('@')[0].lower()),
            'received': last['received'].strftime('%Y-%m-%d %H:%M'),
            'first_received': ms[0]['received'].strftime('%Y-%m-%d'),
            'versions': len(ms),
            'deadline': dl.isoformat() if dl else None,
            'deadline_src': next((x['deadline_src'] for x in reversed(ms) if x['deadline'] == dl), None) if dl else None,
            'period_start': last['period'][0].isoformat() if last['period'] else None,
            'period_end': last['period'][1].isoformat() if last['period'] else None,
            'summary': (hq or ms)[-1].get('summary'), 'action': (hq or ms)[-1].get('action'),
            'changes': (hq or ms)[-1].get('changes'),
            # 給 /sop 頁：系列裡總部最後一封的內文（最後一封常是門市回信）
            'body': (hq or ms)[-1]['body'][:6000] if kind in ('sop', 'event') else None,
            'body_from': f"{(hq or ms)[-1]['sender_name'] or (hq or ms)[-1]['sender']}　"
                         f"{(hq or ms)[-1]['received']:%Y-%m-%d}" if kind in ('sop', 'event') else None,
            'attachments': last['files'][:8],
            'excerpt': last['body'][:600],
            'history': [{'received': x['received'].strftime('%Y-%m-%d %H:%M'), 'subject': x['subject']}
                        for x in ms[:-1]][-10:],
        })
    items.sort(key=lambda x: x['received'], reverse=True)
    return {'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M'), 'days': days,
            'n_mails': len(keep), 'items': items}


# ---------- 門市 SOP／活動頁（hub /sop，所有同事看得到） ----------
SOP_DAYS = 180
# 不給門市看：內部簽核、廣告、員工福利、徵才、成效追蹤、內部行銷製作、人資、預算、客訴、盤商名單
SOP_EXCLUDE = re.compile(
    r'會辦單|簽呈|尚未簽核|Salesforce|Uber|福委|員工|聚餐|徵才|中高齡|銀髮人才|工作坊|AI 創意大賽|'
    r'成效追蹤|DM ?製作|DM ?設計|圖檔製作|校稿|FB ?宣傳|頁面確認|切版|預算|人資|投訴|盤商', re.I)
SOP_EXCLUDE_SENDERS = re.compile(r'星展銀行|內部簽核系統')


def sop_snapshot():
    items = []
    for it in build(SOP_DAYS, False)['items']:
        if it['kind'] not in ('sop', 'event'):
            continue
        if SOP_EXCLUDE.search(it['subject']) or SOP_EXCLUDE_SENDERS.search(it['sender_name'] or ''):
            continue
        items.append({k: it[k] for k in (
            'topic', 'kind', 'subject', 'received', 'first_received', 'versions', 'period_start', 'period_end',
            'summary', 'changes', 'attachments', 'body', 'body_from', 'history')})
    return {'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M'), 'days': SOP_DAYS, 'items': items}


def push(snap, path='/api/mail-digest/snapshot'):
    token = TOKEN_FILE.read_text().strip()
    req = urllib.request.Request(HUB + path,
                                 data=json.dumps(snap, ensure_ascii=False).encode(),
                                 headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())


# ---------- Discord 到期提醒 ----------
REMIND_DAYS = 2   # 剩 0～2 天的交辦才提醒（hub 勾完成的還不會排除，所以窗口放短）


def remind(items):
    """只推交辦（task）：日報、週報、門市回報、公告都不推。"""
    today = date.today()
    due = []
    for it in items:
        if it['kind'] != 'task' or not it.get('deadline'):
            continue
        left = (date.fromisoformat(it['deadline']) - today).days
        if 0 <= left <= REMIND_DAYS:
            due.append((left, it))
    if not due:
        return 0
    due.sort(key=lambda x: (x[0], x[1]['deadline']))
    blocks = []
    for n, (left, it) in enumerate(due, 1):
        subject = re.sub(r'^\s*((RE|FW|FWD|回覆|轉寄)\s*[:：]\s*)+', '', it['subject'], flags=re.I).strip()
        url = re.search(r'https?://[^\s<>」）)\]]+', it.get('excerpt') or '')
        lines = [f'**{n}. {subject}**', f"📅 截止日期：{it['deadline'].replace('-', '/')}（剩餘 **{left}** 天）"]
        if it.get('action'):
            lines.append(f"📝 {it['action']}")
        lines.append(f'🔗 連結：{url.group(0) if url else HUB + "/mail-digest"}')
        blocks.append('\n'.join(lines))
    msgs, cur = [], '🔔 **【待辦事項到期提醒】**'
    for b in blocks:   # Discord 一則上限 2000 字
        if len(cur) + len(b) + 2 > 1900:
            msgs.append(cur)
            cur = ''
        cur += '\n\n' + b
    msgs.append(cur)
    hook = DISCORD_HOOK.read_text().strip()
    for m in msgs:
        # allowed_mentions 空：主旨裡就算有 @everyone／@here 也不會通知全頻道
        req = urllib.request.Request(hook, data=json.dumps({'content': m.strip(),
                                                            'allowed_mentions': {'parse': []}}).encode(),
                                     headers={'Content-Type': 'application/json',
                                              'User-Agent': 'DiscordBot (studioa-mail-digest, 1.0)'})
        urllib.request.urlopen(req, timeout=30).close()
    return len(due)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--days', type=int, default=45)
    ap.add_argument('--no-claude', action='store_true')
    ap.add_argument('--push', action='store_true')
    ap.add_argument('--calendar', action='store_true')
    ap.add_argument('--remind', action='store_true', help='只讀現有 snapshot.json 推 Discord 到期提醒')
    ap.add_argument('--sop', action='store_true', help='另產近 180 天 SOP／活動給 hub /sop（配 --push 才推）')
    a = ap.parse_args()
    if a.remind:
        print('Discord 提醒', remind(json.loads(SNAPSHOT.read_text())['items']), '筆')
        return
    snap = build(a.days, not a.no_claude)
    SNAPSHOT.write_text(json.dumps(snap, ensure_ascii=False, indent=1))
    kinds = {}
    for it in snap['items']:
        kinds[it['kind']] = kinds.get(it['kind'], 0) + 1
    print(f"信件 {snap['n_mails']} 封 → {len(snap['items'])} 個主題 {kinds}")
    # 各步驟互不牽連：hub 掛掉不該擋行事曆，行事曆卡住也不該擋推送；最後有錯再以非 0 結束
    failed = []

    def step(name, fn):
        try:
            fn()
        except Exception as e:
            failed.append(name)
            print(f'  {name}失敗：{type(e).__name__} {e}', file=sys.stderr)

    if a.push:
        step('推送', lambda: print('推送：', push(snap)))
    if a.sop:
        def do_sop():
            sop = sop_snapshot()
            print('SOP／活動', len(sop['items']), '個主題',
                  '→ 推送：' + str(push(sop, '/api/sop/snapshot')) if a.push else '')
        step('SOP', do_sop)
    if a.calendar:
        step('行事曆', lambda: print('行事曆新增', sync_calendar(snap['items']), '筆'))
    if failed:
        sys.exit(1)


if __name__ == '__main__':
    main()
