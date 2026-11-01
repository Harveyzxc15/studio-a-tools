#!/usr/bin/env python3
"""
門市日報信「盤點差異」→ 統一格式（北一區日報信用），並可開 Mail 草稿。

規則抓：從信件內文找每段「盤點差異」，逐行判斷差異類型、抓欄位，
套固定格式（見記憶 inventory_diff_report_format）。缺欄位或認不出類型的
標「待確認」並附原文，不猜。

由 arpedia_daily.py 呼叫 print_region_block()。
"""
import html, os, re, subprocess, time
from email.message import EmailMessage
from email.utils import formataddr

OUT_DIR = os.path.expanduser('~/工具中心/輸出/日報草稿')

# 調撥來源辨識用（13 店＋常見別名）
KNOWN_STORES = ['士林', '微風', '美麗華', '阿波羅', '高島屋', '羅東', '永和', '板橋誠品',
                '新西門', '西門', '花蓮', '板橋遠百', '新莊宏匯', '新店裕隆城', '台大']

CODE = r'(?<!\d)(\d{8})(?!\d)'
DATE = r'(?<!\d)(\d{1,2})/(\d{1,2})(?!\d)'
EMPTY = '（未填）'


def _strip(line):
    line = re.sub(r'^[>\s]+', '', line)            # 轉寄引用符號
    return line.replace('￼', '').strip()


def extract_diffs(body):
    """回傳內文所有「盤點差異」的非無差異行（原文）。"""
    lines = [_strip(l) for l in body.splitlines()]
    out = []
    i = 0
    while i < len(lines):
        m = re.match(r'盤點差異\s*[：:]?\s*(.*)$', lines[i])
        if not m:
            i += 1
            continue
        chunk = [m.group(1)] if m.group(1) else []
        j = i + 1
        while j < len(lines) and j - i <= 15:
            l = lines[j]
            if l.startswith('盤點人員') or l.startswith('盤點差異') or re.match(r'^\d+\s*[\.、．]', l):
                break
            if l:
                chunk.append(l)
            j += 1
        for c in chunk:
            if not re.fullmatch(r'(無差異|無|N/?A)[。.]?', c, re.I):
                out.append(c)
        i = j
    return out


def _date(s):
    m = re.search(DATE, s)
    return f'{int(m.group(1))}/{int(m.group(2))}' if m else None


def _after(label, s, pat=r'(\S+?)'):
    m = re.search(label + r'\s*[：:]\s*' + pat + r'(?=[，,\s]|$)', s)
    return m.group(1) if m else None


def _qty(s, code):
    m = re.search(r'數量\s*[：:]?\s*(\d+)', s) or re.search(code + r'\s*[*xX＊]\s*(\d+)', s)
    return m.group(1) if m else None


def _store(s):
    for st in sorted(KNOWN_STORES, key=len, reverse=True):
        if st in s:
            return st
    m = re.search(r'([一-鿿]{2,4}大學)', s)
    return m.group(1) if m else None


def normalize(raw):
    """原文一行 → (統一格式, 是否完整)。認不出類型回 (None, False)。"""
    s = raw.replace('#', ' #')
    mc = re.search(CODE, s)
    code = mc.group(1) if mc else None
    q = _qty(s, code) if code else None
    tail = _after(r'末四碼', s, r'(\d{4})')
    if not tail:
        m = re.search(r'#\s*(\d{4})(?!\d)', s)
        tail = m.group(1) if m else None
    ship = _after(r'貨運單號', s, r'(\d+)')
    tno = _after(r'調撥單號', s, r'(\w+)')
    sent = _after(r'寄出日期', s, r'(\d{1,2}/\d{1,2})')
    sent = _date(sent) if sent else None

    if '預收已取' in s or '結清待取貨' in s:
        cat = '預收已取' if '預收已取' in s else '結清待取貨'
        sold = _date(re.sub(r'寄出日期.*', '', s))
        f = [code, q, sold, tail]
        txt = f'{cat}：{code or EMPTY}，數量：{q or EMPTY}，銷售日期：{sold or EMPTY}，單據末四碼：{tail or EMPTY}'
    elif '時間差' in s:
        f = [code, q, tail]
        txt = f'時間差銷售：{code or EMPTY}，數量：{q or EMPTY}，單據末四碼：{tail or EMPTY}'
    elif '重包膜' in s or re.search(r'DOA', s, re.I):
        cat = '重包膜' if '重包膜' in s else 'DOA'
        sent = sent or _date(s)                       # 例：「10/19 寄回HQ」沒寫標籤
        f = [code, q, sent, ship]
        txt = f'{cat}：{code or EMPTY}，數量：{q or EMPTY}，寄出日期：{sent or EMPTY}，貨運單號：{ship or EMPTY}'
    elif '校園' in s or ('寄出日期' in s and '調撥單號' not in s and '調撥' in s):   # 範例「台大門市調撥，寄出日期」沒寫校園
        st = _store(s.replace('校園調撥', ''))
        f = [code, q, st, sent, ship]
        src = f'{st}門市調撥' if st else EMPTY
        txt = f'校園調撥：{code or EMPTY}，數量：{q or EMPTY}，{src}，寄出日期：{sent or EMPTY}，貨運單號：{ship or EMPTY}'
    elif '調撥' in s or 'hotmodel' in s.lower():
        if 'hotmodel' in s.lower():
            cat, src = '補 Hotmodel 庫存先收貨', None
            f = [code, q, tno, ship]
            txt = f'{cat}：{code or EMPTY}，數量：{q or EMPTY}，調撥單號：{tno or EMPTY}，貨運單號：{ship or EMPTY}'
        else:
            cat = '門市調撥收貨異常' if '收貨異常' in s else '調撥'
            st = None if '電商' in s else _store(s)
            src = '電商調撥待扣帳' if '電商' in s else (f'{st}門市調撥' if st else None)
            f = [code, q, src, tno, ship]
            txt = f'{cat}：{code or EMPTY}，數量：{q or EMPTY}，{src or EMPTY}，調撥單號：{tno or EMPTY}，貨運單號：{ship or EMPTY}'
    else:
        return None, False
    return txt, all(f)


def build_diff_section(reports, target, codes):
    """回傳 (信件用行列表, 終端機核對用行列表, 待確認筆數)。codes={店名: 代碼}。"""
    mail, check, pending = [], [], 0
    for store, code in codes.items():
        rec = reports.get((target, code))
        if rec is None:
            check.append(f'  {store}：⚠️ 未收到日報信')
            continue
        raws = extract_diffs(rec.get('body', ''))
        if not raws:
            continue
        mail.append(store)
        check.append(f'  {store}')
        for raw in raws:
            txt, ok = normalize(raw)
            if txt is None:
                pending += 1
                mail.append(f'【待確認】{raw}')
                check.append(f'    ⚠️ 認不出類型  原文：{raw}')
            else:
                if not ok:
                    pending += 1
                mail.append(txt if ok else f'【待確認】{txt}')
                check.append(f'    {"✅" if ok else "⚠️ 缺欄位"} {txt}')
                check.append(f'       原文：{raw}')
    return mail, check, pending


def body_text(demo, sales, diff_lines):
    """sales=[本日, 本週, 本月, 累計]。"""
    out = ['Dear 各位長官,', '', '附件為北ㄧ區日報請查收，謝謝。', '', 'ARpedia',
           f'本日 Demo 數：{demo}', f'本日銷售：{sales[0]}', f'本週銷售：{sales[1]}',
           f'本月銷售：{sales[2]}', f'累計銷售：{sales[3]}', '']
    out += ['本日區內盤點：'] + diff_lines if diff_lines else ['本日區內盤點：無差異']
    return '\n'.join(out)


SIGNATURE = ['Bob 尤君暘', 'bob.yu@studioa.com.tw', '', 'STUDIO A 晶實科技股份有限公司',
             '台北市中山區八德路二段260號7樓']


def make_draft(target, text):
    """產生草稿 .eml（收件人空白、副本 rm@、不帶附件），用 Mail 開啟後「再次傳送」轉成可編輯。"""
    m = EmailMessage()
    m['From'] = formataddr(('Bob', 'Bob.Yu@studioa.com.tw'))
    m['Cc'] = formataddr(('Regional Manager', 'rm@studioa.com.tw'))
    m['Subject'] = f'{target} 北ㄧ區 Daily report'
    m['X-Uniform-Type-Identifier'] = 'com.apple.mail-draft'
    m['X-Unsent'] = '1'
    m.set_content(text + '\n\n\n' + '\n'.join(SIGNATURE) + '\n')
    div = lambda l: f'<div>{html.escape(l) if l else "<br>"}</div>'
    m.add_alternative('<html><body dir="auto" style="font-family:\'PingFang TC\';font-size:14px">'
                      + ''.join(div(l) for l in text.split('\n'))
                      + '<br><br><br>' + ''.join(div(l) for l in SIGNATURE) + '</body></html>',
                      subtype='html')
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, f'{target}_北一區日報草稿.eml')
    with open(path, 'wb') as fh:
        fh.write(bytes(m))
    subprocess.run(['open', '-a', 'Mail', path], check=True)
    time.sleep(3)
    # 開啟的 .eml 是唯讀檢視，點「郵件 > 再次傳送」才變成可編輯草稿
    r = subprocess.run(['osascript', '-e', '''
tell application "Mail" to activate
tell application "System Events" to tell process "Mail"
  repeat with mb in menu bar items of menu bar 1
    try
      click menu item "再次傳送" of menu 1 of mb
      return "ok"
    end try
  end repeat
end tell
return "fail"'''], capture_output=True, text=True)
    return path, r.stdout.strip() == 'ok'


def print_region_block(reports, target, yesterday, ar_totals, codes):
    """印盤點差異核對＋日報信內文（可複製），終端機互動時詢問是否開 Mail 草稿。
    ar_totals=[本日, 本週, 本月, 累計] 或 None（EPB 失敗）。"""
    import sys
    from datetime import datetime
    print('【北ㄧ區日報信：區內盤點差異】')
    if not target:
        print('  ⚠️  沒有門市日報信，略過')
        return
    if datetime.strptime(target, '%Y%m%d').date() != yesterday:
        print(f'  ⚠️  門市日報最新日期是 {target}，與 ARpedia 銷售日期 {yesterday} 不同，數字不可直接合用')
    demo = sum(int(r['demo']) for c in codes.values()
               if (r := reports.get((target, c))) and r['demo'] is not None)
    no_demo = [s for s, c in codes.items()
               if (r := reports.get((target, c))) is None or r['demo'] is None]
    mail, check, pending = build_diff_section(reports, target, codes)
    print('\n'.join(check) if check else '  六店皆無差異')
    if no_demo:
        print(f'  ⚠️  Demo 數未計入（未收到信或沒填）：{"、".join(no_demo)}')
    if pending:
        print(f'  ⚠️  有 {pending} 筆【待確認】，寄出前請對原文補齊')
    if ar_totals is None:
        print('  ⚠️  ARpedia 銷售查詢失敗，不產生信件內文')
        return
    text = body_text(demo, ar_totals, mail)
    print(f'\n【{target} 北ㄧ區 Daily report 內文（可複製）】')
    print('-' * 60)
    print(text)
    print('-' * 60)
    if not sys.stdin.isatty():                     # 08:00 排程不互動
        return
    try:
        ans = input('要開 Mail 草稿嗎？(y/N) ').strip().lower()
    except EOFError:
        return
    if ans != 'y':
        return
    path, editable = make_draft(target, text)
    print(f'  草稿檔：{path}')
    if not editable:
        print('  ⚠️  沒辦法自動轉成可編輯：請在 Mail 選單「郵件 > 再次傳送」')
    print('  DSS 截圖請自己貼；收件人空白、副本 rm@')
