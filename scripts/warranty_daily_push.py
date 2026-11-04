#!/usr/bin/env python3
"""每日追蹤主機：跑 warranty_report_all.py → 北一、北二兩張「每日追蹤主機」分頁各截一張 PNG → 推 Discord。

Mac mini 每天 12:00 由 launchd（com.bob.warranty-daily）執行，區間＝本週週日～昨天
（週日跑時 --this-week 會算成上週日～週六整週）。
截圖不開 Excel：直接讀產出的 xlsx（顏色、合併格、欄寬、條件式格式），Python 算出公式值，
排成 HTML 再用 Chrome headless 截圖。圖上的數字就是那份 Excel 的格子。

用法：
  python3 warranty_daily_push.py                 # 本週週日～昨天，推 Discord
  python3 warranty_daily_push.py 11/1 11/3       # 指定區間（測試用）
  python3 warranty_daily_push.py --dry ...       # 只產圖，不推 Discord
"""
import contextlib, html, io, json, math, re, subprocess, sys, tempfile, uuid
import urllib.request
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

sys.path.insert(0, str(Path(__file__).resolve().parent))
import warranty_report_all as W

CHROME  = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
WEBHOOK = Path('~/.config/studioa/mail_digest_discord_webhook.txt').expanduser()   # 跟信件追蹤同頻道
SHEETS  = ['北一-每日追蹤主機', '北二-每日追蹤主機']

# ── 公式求值：只認 warranty_report_all.py 會寫的三種公式，其他一律報錯 ──
RE_SUM  = re.compile(r'^=SUM\(([A-Z]+)(\d+):\1(\d+)\)$')
RE_RATE = re.compile(r'^=IF\(([A-Z]+\d+)>0,N\(([A-Z]+\d+)\)/\1,"—"\)$')
RE_GAP  = re.compile(r'^=IF\(([A-Z]+\d+)<=0,"—",IF\(MAX\(0,CEILING\(ROUND\(\1\*(\d+)/100,6\),1\)-N\(([A-Z]+\d+)\)\)=0,'
                     r'"達標",MAX\(0,CEILING\(ROUND\(\1\*\2/100,6\),1\)-N\(\3\)\)\)\)$')

def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0

def value(ws, ref):
    v = ws[ref].value
    if not (isinstance(v, str) and v.startswith('=')):
        return v
    m = RE_SUM.match(v)
    if m:
        col, r1, r2 = m.group(1), int(m.group(2)), int(m.group(3))
        return sum(_num(value(ws, f'{col}{r}')) for r in range(r1, r2 + 1))
    m = RE_RATE.match(v)
    if m:
        h = _num(value(ws, m.group(1)))
        return _num(value(ws, m.group(2))) / h if h > 0 else '—'
    m = RE_GAP.match(v)
    if m:
        h = _num(value(ws, m.group(1)))
        if h <= 0:
            return '—'
        need = max(0, math.ceil(round(h * int(m.group(2)) / 100, 6)) - _num(value(ws, m.group(3))))
        return '達標' if need == 0 else need
    raise ValueError(f'{ws.title}!{ref} 公式認不得：{v}')

def display(v, number_format):
    if v is None:
        return ''
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if number_format == '0%':
            return f"{(Decimal(repr(v)) * 100).quantize(Decimal('1'), ROUND_HALF_UP)}%"
        return str(int(v)) if float(v).is_integer() else f'{v:.10g}'
    return str(v)

# ── 條件式格式：同樣只認程式會寫的四種 ──
def cf_color(ws, ref, v):
    """回傳條件式格式套出的字色（沒套到回 None）。"""
    hits = []
    for cf in ws.conditional_formatting:
        for rng in str(cf.sqref).split():
            c1, r1, c2, r2 = range_boundaries(rng)
            cell = ws[ref]
            if not (c1 <= cell.column <= c2 and r1 <= cell.row <= r2):
                continue
            top = f'{get_column_letter(c1)}{r1}'
            for rule in cf.rules:
                f = rule.formula[0].replace(top, 'X')
                if f == 'ISTEXT(X)':
                    ok = isinstance(v, str)
                elif f == 'ISNUMBER(X)':
                    ok = isinstance(v, (int, float))
                elif f == 'X="—"':
                    ok = v == '—'
                elif (m := re.fullmatch(r'X<(\d+)/100', f)):
                    ok = isinstance(v, (int, float)) and v < int(m.group(1)) / 100
                else:
                    raise ValueError(f'{ws.title}!{ref} 條件式格式認不得：{rule.formula[0]}')
                if ok:
                    hits.append((rule.priority, rule.dxf.font.color.rgb[-6:], rule.stopIfTrue))
    for _, color, _stop in sorted(hits):
        return color    # 優先序最高的那條；程式寫的規則彼此互斥
    return None

# ── xlsx 分頁 → HTML ──
def _rgb(color):
    rgb = getattr(color, 'rgb', None)
    return rgb[-6:] if isinstance(rgb, str) and len(rgb) >= 6 else None

def sheet_html(ws):
    ncol, nrow = ws.max_column, ws.max_row
    widths  = [round((ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7 + 5) for c in range(1, ncol + 1)]
    heights = [round((ws.row_dimensions[r].height or 15) * 4 / 3) for r in range(1, nrow + 1)]
    span, covered = {}, set()
    for mr in ws.merged_cells.ranges:
        span[(mr.min_row, mr.min_col)] = (mr.max_row - mr.min_row + 1, mr.max_col - mr.min_col + 1)
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    covered.add((r, c))
    rows = []
    for r in range(1, nrow + 1):
        tds = []
        for c in range(1, ncol + 1):
            if (r, c) in covered:
                continue
            cell = ws.cell(r, c)
            ref = cell.coordinate
            v = value(ws, ref)
            style = []
            if cell.fill and cell.fill.fill_type == 'solid' and _rgb(cell.fill.fgColor):
                style.append(f'background:#{_rgb(cell.fill.fgColor)}')
            fg = cf_color(ws, ref, v) or _rgb(cell.font.color) or '000000'
            style.append(f'color:#{fg}')
            style.append(f'font-size:{(cell.font.sz or 10) * 4 / 3:.1f}px')
            if cell.font.bold:
                style.append('font-weight:bold')
            if cell.border and cell.border.left and cell.border.left.style:
                style.append('border:1px solid #BFBFBF')
            rs, cs = span.get((r, c), (1, 1))
            attr = (f' rowspan="{rs}"' if rs > 1 else '') + (f' colspan="{cs}"' if cs > 1 else '')
            text = html.escape(display(v, cell.number_format)).replace('\n', '<br>')
            tds.append(f'<td{attr} style="{";".join(style)}">{text}</td>')
        rows.append(f'<tr style="height:{heights[r-1]}px">{"".join(tds)}</tr>')
    cols = ''.join(f'<col style="width:{w}px">' for w in widths)
    W_, H_ = sum(widths), sum(heights)
    page = f'''<!doctype html><html><head><meta charset="utf-8"><style>
html,body{{margin:0;background:#fff}}
table{{border-collapse:collapse;table-layout:fixed;width:{W_}px;font-family:Arial,"PingFang TC","Heiti TC",sans-serif}}
td{{text-align:center;vertical-align:middle;padding:0 2px;overflow:hidden;line-height:1.25;box-sizing:border-box}}
</style></head><body><table><colgroup>{cols}</colgroup>{"".join(rows)}</table></body></html>'''
    return page, W_, H_

def screenshot(page, w, h, png):
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / 'sheet.html'
        # 長店名（板橋誠品…）會換行把列撐高：先讓 Chrome 排版量出表格實際高度，再用那個高度截圖
        src.write_text(page.replace('</body>', '<script>document.body.dataset.h='
                                    'document.querySelector("table").offsetHeight</script></body>'), encoding='utf-8')
        base = [CHROME, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run',
                f'--user-data-dir={tmp}/profile']
        r = subprocess.run(base + [f'--window-size={w},{h}', '--dump-dom', src.as_uri()],
                           capture_output=True, text=True, timeout=120)
        m = re.search(r'data-h="(\d+)"', r.stdout)
        h = max(h, int(m.group(1))) if m else h + 200
        # 視窗多留 200px 再裁回表格大小：mini 上 Chrome 實際可視高度偶爾比 window-size 小，底部會被切掉、補上重複內容
        r = subprocess.run(base + ['--force-device-scale-factor=2', f'--window-size={w},{h + 200}',
                                   f'--screenshot={png}', src.as_uri()],
                           capture_output=True, text=True, timeout=120)
        if not Path(png).exists():
            raise RuntimeError(f'Chrome 截圖失敗：{r.stderr[-500:]}')
        crop_png_top(png, h * 2)

def crop_png_top(png, height):
    """只留 PNG 最上面 height 列（sips 的 --cropOffset 無效、一律從中間裁，不能用）。
    PNG 每列的濾波只參考上一列，所以截掉下面的列不用重新編碼像素。"""
    import struct, zlib
    data = Path(png).read_bytes()
    pos, chunks = 8, []
    while pos < len(data):
        n, = struct.unpack('>I', data[pos:pos + 4])
        chunks.append((data[pos + 4:pos + 8], data[pos + 8:pos + 8 + n]))
        pos += 12 + n
    ihdr = chunks[0][1]
    w, h, depth, ctype, _, _, interlace = struct.unpack('>IIBBBBB', ihdr)
    if interlace or depth != 8 or h <= height:
        return
    row = 1 + w * {2: 3, 6: 4, 0: 1, 4: 2}[ctype]
    raw = zlib.decompress(b''.join(d for t, d in chunks if t == b'IDAT'))[:row * height]
    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    out = data[:8] + chunk(b'IHDR', struct.pack('>II', w, height) + ihdr[8:])
    out += b''.join(chunk(t, d) for t, d in chunks[1:] if t not in (b'IDAT', b'IEND'))
    out += chunk(b'IDAT', zlib.compress(raw, 9)) + chunk(b'IEND', b'')
    Path(png).write_bytes(out)

def post_discord(content, pngs):
    url = WEBHOOK.read_text().strip()
    boundary = uuid.uuid4().hex
    body = io.BytesIO()
    def part(headers, data):
        body.write(f'--{boundary}\r\n{headers}\r\n\r\n'.encode())
        body.write(data)
        body.write(b'\r\n')
    part('Content-Disposition: form-data; name="payload_json"\r\nContent-Type: application/json',
         json.dumps({'content': content}, ensure_ascii=False).encode())
    for i, p in enumerate(pngs):
        part(f'Content-Disposition: form-data; name="files[{i}]"; filename="{Path(p).name}"\r\nContent-Type: image/png',
             Path(p).read_bytes())
    body.write(f'--{boundary}--\r\n'.encode())
    req = urllib.request.Request(url, data=body.getvalue(), method='POST', headers={
        'Content-Type': f'multipart/form-data; boundary={boundary}', 'User-Agent': 'studioa-warranty-daily/1.0'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        if resp.status >= 300:
            raise RuntimeError(f'Discord 回 {resp.status}')

def main():
    args = sys.argv[1:]
    dry = '--dry' in args
    args = [a for a in args if a != '--dry'] or ['--this-week']
    sys.argv = ['warranty_report_all.py', *args, '--no-open']
    log = io.StringIO()
    class Tee(io.TextIOBase):
        def write(self, s):
            log.write(s); sys.__stdout__.write(s); return len(s)
        def flush(self):
            sys.__stdout__.flush()
    with contextlib.redirect_stdout(Tee()):
        xlsx = W.main()
    out = log.getvalue()
    period = re.search(r'區間：(\S+ ～ \S+)', out).group(1)
    warns = [l.strip() for l in out.splitlines() if l.strip().startswith('⚠️')]

    wb = load_workbook(xlsx)
    # EPB 沒連上（VPN 斷）時查詢靜默回空 → 整張 0：不推圖，改推警告
    hosts = sum(_num(wb[n].cell(r, c).value) for n in SHEETS for r in range(4, wb[n].max_row + 1)
                for c in range(2, 2 + len(W.CATS) * 5, 5))
    if hosts == 0:
        msg = f'⚠️ 每日追蹤主機 {period}：EPB 主機台數全部是 0（公司 VPN 可能沒連上），今天沒推圖'
        print(msg)
        if not dry:
            post_discord(msg, [])
        sys.exit(1)
    pngs = []
    for name in SHEETS:
        page, w, h = sheet_html(wb[name])
        png = Path(xlsx).with_name(f'{Path(xlsx).stem}_{name[:2]}.png')
        png.unlink(missing_ok=True)
        screenshot(page, w, h, png)
        pngs.append(png)
        print(f'🖼️  {png}')
    if dry:
        print('（--dry：不推 Discord）')
        return
    # 一則訊息一張圖：同一則附兩張，Discord 會排成拼貼並把圖裁成方形縮圖
    for name, png in zip(SHEETS, pngs):
        region = name[:2] + '區'
        content = '\n'.join([f'📊 每日追蹤主機 保固搭售率　{region}　區間 {period}', *(warns if region == '北一區' else [])])
        post_discord(content, [png])
    print('✅ 已推 Discord')

if __name__ == '__main__':
    main()
