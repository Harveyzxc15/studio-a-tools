"""apple_duo_stock 模擬測試（假 Apple 回應／Discord／時間，不連網）：python3 test_apple_duo_stock.py"""
import importlib.util
import io
import json
import sys
import tempfile
import urllib.error
from pathlib import Path

spec = importlib.util.spec_from_file_location("duo", Path(__file__).with_name("apple_duo_stock.py"))
a = importlib.util.module_from_spec(spec)
spec.loader.exec_module(a)
REAL_DISCORD = a.discord

W, B, S = "MK2H4ZP/A", "MK2J4ZP/A", "信義 A13"


def resp(stores):
    """stores: {店名: {型號: pickupDisplay}} → 假的 Apple JSON"""
    return {"body": {"stores": [
        {"storeName": n, "partsAvailability": {p: {"pickupDisplay": d, "pickupSearchQuote": "今天" if d == "available" else "目前無法提供"}
                                               for p, d in parts.items()}} for n, parts in stores.items()]}}


def both(w="unavailable", b="unavailable"):
    return resp({S: {W: w, B: b}, "台北 101": {W: "unavailable", B: "unavailable"}})


import time as _t


class FakeTime:
    """可控時間：預設 2026-10-26 08:00（心跳時間前）"""
    def __init__(self):
        self.now = _t.mktime((2026, 10, 26, 8, 0, 0, 0, 0, -1))

    def time(self):
        return self.now

    def localtime(self, t=None):
        return _t.localtime(self.now if t is None else t)

    def strftime(self, fmt, t=None):
        return _t.strftime(fmt, self.localtime() if t is None else t)

    def at(self, day, hour, minute=0):
        self.now = _t.mktime((2026, 10, day, hour, minute, 0, 0, 0, -1))


class Env:
    def __init__(self):
        self.clock = a.time = FakeTime()
        d = Path(tempfile.mkdtemp())
        a.STATE, a.LOG = d / "state.json", d / "log.txt"
        self.sent, self.discord_down = [], False

        def fake_discord(msg):
            if self.discord_down:
                raise urllib.error.URLError("discord down")
            self.sent.append(msg)
        a.discord = fake_discord

    def run(self, data):
        def fake_urlopen(req, timeout=None):
            if isinstance(data, Exception):
                raise data
            return io.BytesIO(json.dumps(data).encode())
        a.urllib.request.urlopen = fake_urlopen
        a.main()

    def stocked(self):
        return [m for m in self.sent if "可在" in m]

    def alerts(self):
        return [m for m in self.sent if "⚠️" in m]


fails = 0


def check(name, cond):
    global fails
    print(("✅" if cond else "❌"), name)
    fails += not cond


# 1. 基本：缺貨→有貨通知一次、持續有貨不重複、售完再補貨再通知
e = Env()
for d in [both(), both(b="available"), both(b="available"), both(), both(b="available")]:
    e.run(d)
check("基本轉換只在缺貨→有貨通知", len(e.stocked()) == 2)

# 2. 型號從回應消失（下架/改代碼）要算失敗並告警，不能當成沒貨
e = Env()
for _ in range(a.FAIL_ALERT):
    e.run(resp({S: {W: "unavailable"}, "台北 101": {W: "unavailable"}}))
check("型號消失會累計失敗並告警", len(e.alerts()) == 1)

# 3. Apple 回空門市清單也要算失敗
e = Env()
for _ in range(a.FAIL_ALERT):
    e.run(resp({}))
check("空門市清單會告警", len(e.alerts()) == 1)

# 4. 狀態檔損毀不能讓程式永久停擺
e = Env()
a.STATE.write_text('{"stock": {"信義')
try:
    e.run(both(b="available"))
    check("狀態檔損毀後自動復原並照常通知", len(e.stocked()) == 1)
except Exception as ex:
    check(f"狀態檔損毀後自動復原（拋出 {type(ex).__name__}）", False)

# 5. Discord 暫時發不出去：下一輪要重試，且同輪其他品項狀態要存下來
e = Env()
e.run(both())
e.discord_down = True
try:
    e.run(both(w="available", b="available"))
    no_crash = True
except Exception:
    no_crash = False
check("Discord 失敗不讓程式崩潰", no_crash)
e.discord_down = False
e.run(both(w="available", b="available"))
check("Discord 恢復後補發兩則有貨通知", len(e.stocked()) == 2)

# 6. 門市暫時從回應消失、再出現時有貨，要通知
e = Env()
e.run(both(b="available"))                      # 通知 1
e.run(resp({"台北 101": {W: "unavailable", B: "unavailable"}, S: {W: "unavailable", B: "unavailable"}}))
e.run(resp({"台北 101": {W: "unavailable", B: "unavailable"}}))  # 信義消失（型號仍在 101 → 不算失敗）
e.run(both(b="available"))                      # 信義回來且有貨 → 通知 2
check("售完→門市消失→回來有貨會再通知", len(e.stocked()) == 2)

# 7. HTTP 541 被擋算失敗
e = Env()
for _ in range(a.FAIL_ALERT):
    e.run(urllib.error.HTTPError("u", 541, "blocked", {}, None))
check("HTTP 541 累計告警", len(e.alerts()) == 1)

# 8. 恢復通知：達告警門檻後恢復才發 ✅；沒到門檻的短暫失敗不發
e = Env()
for _ in range(a.FAIL_ALERT):
    e.run(urllib.error.URLError("x"))
e.run(both())
e.run(both())
recov = [m for m in e.sent if m.startswith("✅")]
check("達門檻後恢復發一則 ✅", len(recov) == 1 and f"{a.FAIL_ALERT} 次" in recov[0])
e = Env()
for _ in range(a.FAIL_ALERT - 1):
    e.run(urllib.error.URLError("x"))
e.run(both())
check("未達門檻的短暫失敗不發告警也不發恢復", not e.alerts() and not [m for m in e.sent if m.startswith("✅")])

# 9. 售完通知：帶有貨時長、不 @everyone、只發一次
e = Env()
e.run(both())
e.run(both(b="available"))
e.clock.at(26, 8, 15)
e.run(both())
e.clock.at(26, 8, 20)
e.run(both())
sold = [m for m in e.sent if "已售完" in m]
check("售完發一次且時長 15 分鐘", len(sold) == 1 and "約 15 分鐘" in sold[0] and "@everyone" not in sold[0])

# 10. 有貨通知沒發出去時，之後售完不能發「已售完」
e = Env()
e.discord_down = True
e.run(both(b="available"))
e.discord_down = False
e.run(both())
check("沒通知過有貨就不發售完", not [m for m in e.sent if "已售完" in m])

# 11. 心跳：9 點前不發、9 點後每天一次、隔天再發、發失敗會重試
e = Env()
e.run(both())                       # 08:00
hb = lambda: [m for m in e.sent if m.startswith("💓")]
check("9 點前不發心跳", not hb())
e.clock.at(26, 9, 0); e.run(both(b="available"))
e.clock.at(26, 9, 5); e.run(both(b="available"))
check("9 點後當天只發一次，且列出庫存", len(hb()) == 1 and "夜空色：有貨" in hb()[0] and len(hb()[0].splitlines()) == 5)
e.clock.at(27, 9, 0); e.discord_down = True; e.run(both())
e.clock.at(27, 9, 5); e.discord_down = False; e.run(both())
check("隔天再發、發失敗下輪重試", len(hb()) == 2)

# 12. 舊版狀態檔（stock 值是 false）能直接沿用
e = Env()
a.STATE.write_text(json.dumps({"stock": {f"{S}|{B}": False}, "fails": 0}))
e.run(both(b="available"))
check("沿用舊版狀態檔", len(e.stocked()) == 1)

# 13. repo 是公開的：程式碼不能寫死 webhook；設定檔不見時只記 log、不崩潰
check("程式碼不含 webhook 網址", "discord.com/api/webhooks" not in Path(spec.origin).read_text())
e = Env()
a.discord = REAL_DISCORD
a.WEBHOOK_FILE = a.STATE.with_name("no_such_webhook")
check("webhook 設定檔不見時 send() 回 False", a.send("x") is False and "Discord 發送失敗" in a.LOG.read_text())

print("全部通過" if not fails else f"{fails} 項失敗")
sys.exit(1 if fails else 0)
