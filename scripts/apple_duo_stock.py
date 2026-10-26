#!/usr/bin/env python3
"""Apple 台灣直營店 iPhone Duo 店取庫存追蹤 → Discord 通知

launchd 每 5 分鐘執行一次。缺貨→有貨才通知（@everyone），售完發一則（不 @），之後自動重新武裝。
連續 3 次查詢失敗發告警、恢復時再通知；每天 9 點後第一次查詢發心跳。
只查庫存、不自動下單，結帳請手動用 Apple Pay。
"""
import json
import os
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

PARTS = {
    "MK2H4ZP/A": "iPhone Duo 1TB 星光白色",
    "MK2J4ZP/A": "iPhone Duo 1TB 夜空色",
}
LOCATION = "110"  # 台北 101、信義 A13 都會回（台灣只有這兩家）
# /tw/shop/product/* 是 Apple Store App 的通用連結：iPhone 點了直接開 App 並選好型號顏色
PRODUCT_URL = "https://www.apple.com/tw/shop/product/{part}"
UA ="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15"

DIR = Path(__file__).resolve().parent
STATE = DIR / ".apple_duo_state.json"
LOG = DIR / "apple_duo_stock.log"
WEBHOOK_FILE = DIR / ".apple_duo_webhook"  # 一行 Discord webhook 網址；含帳密，不進 git（repo 是公開的）
FAIL_ALERT = 3
HEARTBEAT_HOUR = 9


def log(msg):
    with LOG.open("a") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}\n")


def discord(content):
    body = json.dumps({"content": content, "allowed_mentions": {"parse": ["everyone"]}}).encode()
    req = urllib.request.Request(WEBHOOK_FILE.read_text().strip(), body,{"Content-Type": "application/json", "User-Agent": UA})
    urllib.request.urlopen(req, timeout=15)


def send(content):
    """發 Discord；失敗只記 log 回傳 False，不讓整輪中斷"""
    try:
        discord(content)
        return True
    except Exception as e:
        log(f"Discord 發送失敗：{e}")
        return False


def load_state():
    try:
        return json.loads(STATE.read_text())
    except (FileNotFoundError, ValueError):  # 不存在或寫到一半損毀 → 重來
        return {"stock": {}, "fails": 0}


def save_state(state):
    tmp = STATE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False))
    os.replace(tmp, STATE)  # 原子替換，斷電也不會留下半個檔


def fetch():
    """回傳 {(store, part): (available, quote)}"""
    q = {"pl": "true", "location": LOCATION}
    q.update({f"parts.{i}": p for i, p in enumerate(PARTS)})
    url = "https://www.apple.com/tw/shop/retail/pickup-message?" + urllib.parse.urlencode(q)
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    data = json.load(urllib.request.urlopen(req, timeout=20))
    out = {}
    for s in data["body"]["stores"]:
        for part, a in s["partsAvailability"].items():
            out[(s["storeName"], part)] = (a["pickupDisplay"] == "available", a["pickupSearchQuote"])
    # 型號下架或代碼錯時 Apple 仍回 200，只是沒有該型號；不擋的話會被當成「沒貨」永遠沉默
    missing = [p for p in PARTS if p not in {part for _, part in out}]
    if missing:
        raise ValueError(f"回應中沒有型號 {', '.join(missing)}（下架、改代碼或 Apple 改格式）")
    return out


def main():
    state = load_state()
    try:
        result = fetch()
    except Exception as e:
        state["fails"] += 1
        log(f"查詢失敗 {state['fails']}：{e}")
        if state["fails"] == FAIL_ALERT:
            send(f"⚠️ Apple Duo 庫存查詢連續失敗 {FAIL_ALERT} 次：{e}")
        save_state(state)
        return
    if state["fails"] >= FAIL_ALERT:
        send(f"✅ Apple Duo 庫存查詢恢復正常（先前連續失敗 {state['fails']} 次）")
    state["fails"] = 0

    now = time.time()
    for (store, part), (avail, quote) in result.items():
        key = f"{store}|{part}"
        since = state["stock"].get(key)  # 有貨通知發出的時間；False＝沒貨
        name = PARTS.get(part, part)
        if avail and not since:
            if send(f"@everyone 🍎 **{name}** 可在 **Apple {store}** 店取！（{quote}）\n"
                    f"👉 {PRODUCT_URL.format(part=part)}"):
                log(f"有貨通知 {store} {part} {quote}")
                since = now
            # 沒發出去就維持沒貨，下一輪重試
        elif not avail and since:
            send(f"📦 {name} 在 Apple {store} 已售完（有貨約 {round((now - since) / 60)} 分鐘）")
            log(f"售完 {store} {part}")
            since = False
        state["stock"][key] = since or False

    today = time.strftime("%Y-%m-%d")
    if time.localtime().tm_hour >= HEARTBEAT_HOUR and state.get("heartbeat") != today:
        lines = [f"{store}｜{PARTS.get(part, part)}：{'有貨' if avail else '缺貨'}"
                 for (store, part), (avail, _) in result.items()]
        if send("💓 Apple Duo 庫存追蹤運作中（每 5 分鐘查一次）\n" + "\n".join(lines)):
            state["heartbeat"] = today
    save_state(state)
    if "-v" in sys.argv:
        for (store, part), (avail, quote) in result.items():
            print(store, PARTS.get(part, part), "有貨" if avail else "缺貨", quote)


if __name__ == "__main__":
    main()
