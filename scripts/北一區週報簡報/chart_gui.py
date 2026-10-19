"""改 Keynote 原生圖表的資料（保留原圖表所有樣式）。

Keynote 的 AppleScript 讀不到也改不了圖表資料，只能模擬人工操作「圖表資料編輯器」：
選圖表 → 格式 › 圖表 › 顯示圖表資料編輯器 → 右鍵「全選」→ 刪除 → 貼上新類別列與數列 →
再「全選」拷貝回來逐格比對。比對不符就丟例外，由呼叫端改用圖片取代。
執行期間會操作畫面（約 10 秒／張），不可同時動鍵盤滑鼠；剪貼簿會被暫用、結束後還原文字內容。
"""
import subprocess
import time

from AppKit import NSPasteboard, NSPasteboardTypeString

PROC = 'tell application "System Events" to tell (first process whose bundle identifier is "com.apple.Keynote")'
# 「顯示／隱藏圖表資料編輯器」：名稱是上次選單驗證時的狀態、不一定即時，所以用結尾比對，點完再確認實際開關
CHART_MENU = ('(first menu item of menu 1 of menu item "圖表" of menu "格式" of menu bar item "格式" '
              'of menu bar 1 whose name ends with "圖表資料編輯器")')
HAS_EDITOR = '(exists (first window whose subrole is "AXDialog"))'
# 圖表資料編輯器內的位置（相對編輯器視窗左上角，單位點）
STRIP = (150, 46)       # 第一個資料直欄上方的空白列（右鍵選單有「全選」）
FIRST_LABEL = (150, 69)  # 第一個類別名稱格
FIRST_NAME = (61, 96)    # 第一個數列名稱格


class ChartGuiError(Exception):
    pass


def _as(script):
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=60)
    if r.returncode != 0:
        raise ChartGuiError(r.stderr.strip())
    return r.stdout.strip()


_MOUSE = '''ObjC.import('CoreGraphics');
function run(a) {
  const p = $.CGPointMake(+a[0], +a[1]), right = a[2] === 'r';
  $.CGEventPost($.kCGHIDEventTap, $.CGEventCreateMouseEvent($(), $.kCGEventMouseMoved, p, 0));
  delay(0.2);
  const ts = right ? [$.kCGEventRightMouseDown, $.kCGEventRightMouseUp] : [$.kCGEventLeftMouseDown, $.kCGEventLeftMouseUp];
  for (const t of ts) {
    $.CGEventPost($.kCGHIDEventTap, $.CGEventCreateMouseEvent($(), t, p, right ? $.kCGMouseButtonRight : $.kCGMouseButtonLeft));
    delay(0.05);
  }
}'''


def _click(x, y, right=False):
    subprocess.run(['osascript', '-l', 'JavaScript', '-e', _MOUSE, str(x), str(y), 'r' if right else 'l'],
                   check=True, capture_output=True, timeout=30)
    time.sleep(0.4)


def _keys(*codes):
    _as(PROC + '\n' + '\n'.join(f'key code {c}\ndelay 0.25' for c in codes) + '\nend tell')


def _cmd(ch):
    _as(f'{PROC} to keystroke "{ch}" using command down')
    time.sleep(0.8)


def _clip_get():
    return NSPasteboard.generalPasteboard().stringForType_(NSPasteboardTypeString)


def _clip_set(s):
    pb = NSPasteboard.generalPasteboard()
    pb.clearContents()
    pb.setString_forType_(s, NSPasteboardTypeString)


def _editor_origin():
    out = _as(PROC + '''
repeat with w in windows
  if subrole of w is "AXDialog" then
    set v to scroll area 1 of w
    repeat with b in scroll bars of v
      set value of value indicator 1 of b to 0
    end repeat
    set p to position of w
    return ((item 1 of p) as text) & "," & ((item 2 of p) as text)
  end if
end repeat
return ""
end tell''')
    if not out:
        raise ChartGuiError('圖表資料編輯器沒有打開')
    x, y = map(int, out.split(','))
    return x, y


def _editor(want_open):
    """打開／關閉圖表資料編輯器，以實際視窗為準。"""
    for _ in range(3):
        if (_as(f'{PROC} to return {HAS_EDITOR}') == 'true') == want_open:
            return
        _as(f'{PROC} to click {CHART_MENU}')
        time.sleep(1.2)
    raise ChartGuiError('圖表資料編輯器無法' + ('打開' if want_open else '關閉'))


def _select_all(ox, oy):
    _click(ox + STRIP[0], oy + STRIP[1])        # 先讓編輯器取得焦點，右鍵選單才會出來
    _click(ox + STRIP[0], oy + STRIP[1], right=True)
    time.sleep(0.3)
    _keys(125, 36)          # 選單第一項「全選」


def _copy_all(ox, oy):
    _clip_set('')
    _select_all(ox, oy)
    _cmd('c')
    return [ln.split('\t') for ln in (_clip_get() or '').strip('\n').split('\n')]


def update_chart(doc_name, slide_no, labels, rows):
    """doc_name：Keynote 文件名稱；slide_no：第幾頁（1 起算，含略過頁）；
    labels：類別（x 軸）；rows：每個數列的數值 list，順序同原圖表（數列名稱沿用原圖表）。"""
    saved = _clip_get()
    try:
        _as(f'''tell application id "com.apple.Keynote"
activate
set d to document "{doc_name}"
set current slide of d to slide {slide_no} of d
end tell
{PROC}
set frontmost to true
set position of (first window whose name is "{doc_name}") to {{0, 25}}
end tell''')
        time.sleep(0.8)
        # Tab 依序選取頁面上的物件，直到選到圖表（圖表選單可用）為止
        ok = _as(PROC + f'''
repeat with i from 1 to 20
  if enabled of {CHART_MENU} then return "ok"
  key code 48
  delay 0.3
end repeat
return ""
end tell''')
        if not ok:
            raise ChartGuiError(f'第 {slide_no} 頁選不到圖表')
        _editor(True)
        ox, oy = _editor_origin()

        names = [r[0] for r in _copy_all(ox, oy)[1:]]
        if len(names) != len(rows):
            raise ChartGuiError(f'第 {slide_no} 頁圖表有 {len(names)} 個數列，預期 {len(rows)} 個：{names}')
        rows = list(zip(names, rows))
        _select_all(ox, oy)
        _keys(51)               # 刪除全部資料（圖表樣式保留）
        time.sleep(0.5)
        _clip_set('\t'.join(labels))
        _click(ox + FIRST_LABEL[0], oy + FIRST_LABEL[1])
        _cmd('v')
        _clip_set('\n'.join('\t'.join([n] + [str(v) for v in vals]) for n, vals in rows))
        _click(ox + FIRST_NAME[0], oy + FIRST_NAME[1])
        _cmd('v')

        # 驗證：全選拷貝回來逐格比對
        got = _copy_all(ox, oy)
        want = [[''] + list(labels)] + [[n] + [str(v) for v in vals] for n, vals in rows]
        norm = lambda g: [[c.strip().replace(',', '') for c in r] for r in g]
        if norm(got) != norm(want):
            raise ChartGuiError(f'第 {slide_no} 頁圖表資料比對不符：{got[:3]}')

        _editor(False)
    finally:
        if saved is not None:
            _clip_set(saved)
