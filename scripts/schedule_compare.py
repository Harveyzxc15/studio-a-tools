#!/usr/bin/env python3
"""門市日報自填人數 vs 排班系統班表 的核對（北一區）。

日報的「本日銷售人數／本日休假人數」是門市自己填的，
這裡拿 schedule-system.fly.dev 的班表重算一份來對，抓填錯的店。

⚠️ 排班系統帳號只看得到北一 6 店，北二區無法核對。
"""
import os
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / '排班系統'))

from daily_headcount import CODES, pad, read_headcount, w

# 門市名 → 排班系統 store_id（名稱以排班系統為準：大葉高島屋＝日報的「高島屋」）
STORE_IDS = {'士林': 1, '微風': 3, '阿波羅': 4,
             '美麗華': 5, '高島屋': 6, '羅東': 7}

# 班表班別 → 休假。其餘有排班的格子一律算上班（含 開會／課程／支援）。
# 空白格（沒排班別也沒排休）門市是算在休假人數裡的，用 '未排班' 併入休假。
OFF_SHIFTS = {'排休', '必休', '休', '特休', '病假', '事假', '喪假', '其他假別', '公出',
              '未排班'}

# 日報不列入的人：羅東維修人員（門市自己就不算在銷售/休假人數裡）
EXCLUDE_ROLES = {'羅東': {'維修'}}


def _has_password():
    """沒密碼就別跑，免得 getpass 卡住日報。"""
    if os.environ.get('SCHED_PWD'):
        return True
    return subprocess.run(
        ['security', 'find-generic-password', '-s', 'schedule-system',
         '-a', os.environ.get('SCHED_ACCOUNT', 'SA1469'), '-w'],
        capture_output=True).returncode == 0


def fetch_day(date_iso):
    """回傳 {門市: [(姓名, 職務, 班別), ...]}，含期別內全部人員；
    該日格子空白者班別記為 '未排班'（算休假）。"""
    import schedule_api

    c = schedule_api.Client()
    out = {}
    for store, sid in STORE_IDS.items():
        c.switch_store(sid)
        periods = [p for p in c.periods()
                   if p['start_date'] <= date_iso <= p['end_date']]
        if not periods:
            out[store] = None                      # 該日沒有對應期別
            continue
        sched = c.schedule(periods[0]['id'])['schedule']
        rows = []
        for e in sched.values():
            role = e.get('role', '')
            if role in EXCLUDE_ROLES.get(store, ()):
                continue
            shift = ((e.get('days') or {}).get(date_iso) or {}).get('shift')
            rows.append((e.get('name', ''), role, shift or '未排班'))
        out[store] = rows
    return out


def print_schedule_compare(reports, target, region='北一區'):
    """印出日報自填人數 vs 班表人數；對不上的店展開人員明細。"""
    if region != '北一區':
        return
    if not reports or not target:
        return
    if not _has_password():
        print("【人力核對：日報自填 vs 排班系統】")
        print("  ⚠️  Keychain 沒有排班系統密碼，略過核對")
        print("      設定一次即可：security add-generic-password "
              "-s schedule-system -a SA1469 -w")
        return

    date_iso = f"{target[:4]}-{target[4:6]}-{target[6:]}"
    print(f"【人力核對：日報自填 vs 排班系統（{region} {date_iso}）】")
    try:
        sched = fetch_day(date_iso)
    except Exception as e:
        print(f"  ⚠️  排班系統讀取失敗：{e}")
        return

    header = (pad('門市', 12) + pad('日報上班', 10, '>') + pad('班表上班', 10, '>')
              + pad('差異', 8, '>') + pad('日報休假', 10, '>')
              + pad('班表休假', 10, '>') + pad('差異', 8, '>'))
    print(header)
    print('-' * w(header))

    def diff(m, s):
        if not isinstance(m, (int, float)):
            return '--'
        d = int(m) - s
        return '一致' if d == 0 else f"{d:+d}"

    details = []
    checked = 0
    for store in CODES[region]:
        rec = reports.get((target, CODES[region][store]))
        rows = sched.get(store)
        if rec is None:
            print(pad(store, 12) + pad('未收到', 10, '>'))
            continue
        if rows is None:
            print(pad(store, 12) + pad('', 10, '>') + pad('無班表', 10, '>'))
            continue
        checked += 1
        m_on, m_off = read_headcount(rec['xlsx'])
        s_on = sum(1 for _, _, sh in rows if sh not in OFF_SHIFTS)
        s_off = len(rows) - s_on
        print(pad(store, 12)
              + pad(m_on if isinstance(m_on, (int, float)) else '異常', 10, '>')
              + pad(s_on, 10, '>') + pad(diff(m_on, s_on), 8, '>')
              + pad(m_off if isinstance(m_off, (int, float)) else '異常', 10, '>')
              + pad(s_off, 10, '>') + pad(diff(m_off, s_off), 8, '>'))
        if diff(m_on, s_on) != '一致' or diff(m_off, s_off) != '一致':
            details.append((store, rows))

    if not details:
        print(f"\n  ✅ 已核對 {checked} 店，上班／休假人數與班表全數相符")
        return

    print("\n  ── 有差異的門市，當日班表明細 ──")
    for store, rows in details:
        print(f"\n  ＜{store}＞")
        for name, role, shift in sorted(rows, key=lambda r: r[2] in OFF_SHIFTS):
            mark = '休' if shift in OFF_SHIFTS else '上'
            print(f"    [{mark}] {pad(name, 12)}{pad(role, 8)}{shift}")
    print("\n  ※ 班表口徑：排休／必休／休／特休／病假／事假／喪假／其他假別／公出／未排班（空白格）"
          "＝休假，其餘有排班＝上班（開會、課程、支援算上班）。羅東維修人員不列入。")


if __name__ == '__main__':
    from daily_headcount import collect_reports
    reps = collect_reports(since_days=4, progress=False)
    tgt = max(t for t, _ in reps) if reps else None
    print_schedule_compare(reps, tgt)
