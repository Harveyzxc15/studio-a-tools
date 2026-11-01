#!/usr/bin/env python3
"""schedule-system.fly.dev 的 API client。

密碼只從 macOS Keychain 或環境變數 SCHED_PWD 讀，不落檔：
  security add-generic-password -s schedule-system -a SA1469 -w
"""
import getpass
import os
import subprocess
import sys

import requests

BASE = "https://schedule-system.fly.dev"
ACCOUNT = os.environ.get("SCHED_ACCOUNT", "SA1469")


def _password():
    if os.environ.get("SCHED_PWD"):
        return os.environ["SCHED_PWD"]
    try:
        return subprocess.run(
            ["security", "find-generic-password", "-s", "schedule-system",
             "-a", ACCOUNT, "-w"],
            capture_output=True, text=True, check=True,
        ).stdout.strip()
    except subprocess.CalledProcessError:
        return getpass.getpass(f"{ACCOUNT} 密碼：")


class Client:
    def __init__(self):
        self.s = requests.Session()
        r = self.s.post(f"{BASE}/api/auth/login",
                        json={"account": ACCOUNT, "email": ACCOUNT,
                              "password": _password()}, timeout=30).json()
        if not r.get("success"):
            sys.exit(f"登入失敗：{r}")
        self.user = r["user"]
        self.stores = self.user.get("stores") or []

    def _get(self, path):
        r = self.s.get(f"{BASE}/api{path}", timeout=60)
        if r.status_code == 401:
            sys.exit("session 失效")
        return r.json()

    def switch_store(self, store_id):
        r = self.s.post(f"{BASE}/api/auth/switch-store",
                        json={"store_id": store_id}, timeout=30).json()
        if not r.get("success"):
            raise RuntimeError(f"切店失敗 {store_id}: {r}")
        self.user = r.get("user") or self.user
        return self.user

    def periods(self):
        r = self._get("/periods")
        return r if isinstance(r, list) else r.get("periods", [])

    def shifts(self):
        """本店有效班別定義 {班別名: {...}}。

        `shifts` 已含「全域啟用 + 門市自訂」的完整清單，是班表格子實際存的名稱；
        `custom_shifts` 只是門市自訂那批（多為封存舊班別），不能拿來當主檔。
        門市可用 shift_overrides 覆寫全域班別的工時／時段，這裡先套用再回傳。
        `alias` 是顯示別名（全07 → 顯示「全班」），僅供輸出用，不是格子的值。
        """
        r = self._get("/store-shifts")
        overrides = r.get("shift_overrides") or {}
        aliases = r.get("shift_aliases") or {}
        out = {}
        for x in r.get("shifts", []):
            sh = dict(x)
            sh.update(overrides.get(sh["name"], {}))
            sh["alias"] = aliases.get(sh["name"]) or sh.get("alias") or ""
            sh["is_global"] = bool(x.get("global_id"))
            out[sh["name"]] = sh
        return out

    def schedule(self, period_id):
        return self._get(f"/schedule/{period_id}")
