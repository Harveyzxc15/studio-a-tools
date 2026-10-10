#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
目標分配 — 把 EPB 實績推到 sa-north-hub
======================================
EPB 只走公司 VPN，雲端打不到，所以在本機查完之後把結果 POST 上去，
同事不必連 VPN 也看得到，這台 Mac 關機也不影響。

查詢邏輯完全沿用 目標分配_填業績.py（同一組 SQL、同一份 SACare 價目表），
這裡只做「查完 → 組 JSON → 送出」，不重寫任何口徑。

【用法】
    python3 目標分配_推雲端.py 202610      # 要訂的月份；會推它的去年同期、上月、前 12 個月高單價客群指標
    python3 目標分配_推雲端.py --poll      # launchd 每 30 秒：跟雲端報到（同步狀態燈）、日報缺天就補推；
                                           #   目標訂定頁有人按「刷新 EPB」才重推、
                                           #   按「上傳到內網」才上傳內網目標網站並回查核對
    HUB_URL=http://127.0.0.1:5000 python3 目標分配_推雲端.py 202610   # 測本機

【Token】環境變數 EPB_PUSH_TOKEN，或 ~/.config/studioa/epb_push_token
======================================
"""
import io
import json
import os
import sys
from datetime import datetime
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import 目標分配_填業績 as F

HUB_URL = os.environ.get('HUB_URL', 'https://sa-north-hub.fly.dev')
TOKEN_FILE = os.path.expanduser('~/.config/studioa/epb_push_token')
HUB_REPO = os.path.expanduser('~/sa_north_hub')
# 內網目標網站（公司網路，要連 VPN）。上傳沒有另外認證；查詢 action=query 不給 store_id 會回全部店
INTRANET_URL = os.environ.get('INTRANET_URL', 'http://192.168.1.186/dssAutoMail/targetMapHandle.php')

# 台數欄名：本機腳本的鍵 → 雲端資料表的欄
QTY_MAP = [('cpu', 'q_cpu'), ('mini', 'q_mini'), ('ipad', 'q_ipad'),
           ('ipadmini', 'q_ipad_mini'), ('ipadpro', 'q_ipad_pro'),
           ('iphone', 'q_iphone'), ('watch', 'q_watch')]


def token():
    value = os.environ.get('EPB_PUSH_TOKEN')
    if not value and os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, encoding='utf-8') as handle:
            value = handle.read().strip()
    if not value:
        sys.exit(f'找不到 token：設環境變數 EPB_PUSH_TOKEN 或寫進 {TOKEN_FILE}')
    return value


STANDBY = os.environ.get('SYNC_STANDBY') == '1'     # 備援機（筆電）的 launchd 設這個；主機（Mac mini）不設
STANDBY_WAIT = 3 * 60
STANDBY_STATE = os.path.expanduser('~/.config/studioa/standby_seen.json')


def standby_ready(kind, keys):
    """備援機：同一件待辦掛滿 3 分鐘、主機都沒處理才接手；主機一律直接做。

    記下每件待辦第一次看到的時間；已經不在待辦裡的就忘掉，下次同一件再出現重新起算。
    """
    if not STANDBY:
        return list(keys)
    try:
        with open(STANDBY_STATE, encoding='utf-8') as handle:
            state = json.load(handle)
    except (OSError, ValueError):
        state = {}
    now = datetime.now().timestamp()
    seen = {key: state.get(kind, {}).get(key, now) for key in keys}
    state[kind] = seen
    with open(STANDBY_STATE, 'w', encoding='utf-8') as handle:
        json.dump(state, handle)
    return [key for key in keys if now - seen[key] >= STANDBY_WAIT]


def store_ids():
    """店名 → store_id，直接用 sa-north-hub 的 stores.py，避免兩邊各存一份會走鐘。"""
    sys.path.insert(0, HUB_REPO)
    try:
        from stores import STORES
    except ImportError:
        sys.exit(f'找不到 {HUB_REPO}/stores.py，無法對應 store_id')
    alias = {'大葉高島屋': '高島屋'}
    return {alias.get(name, name): store_id for store_id, _, name, _ in STORES}


def rows_for(data, sacare, ids):
    rows = []
    for store, values in data.items():
        row = {'store_id': ids[store],
               'revenue': values['totsale'],          # 不含 SACare（EPB 的 SACare 金額是 0）
               'tpp': values['tpp'],
               'acc': values['acc'],
               'sacare': sacare.get(store, 0)}
        row.update({column: values[key] for key, column in QTY_MAP})
        rows.append(row)
    return rows


def query_high_end(target_month, locs):
    """高單價客群指標（Duo 這類高單價新機的分配參考），涵蓋目標月前 12 個月＋上個月。

    口徑與 2026-10 那次 Duo 分析相同：≥6 萬＝iPhone 單台實收，≥10 萬／≥7 萬＝單筆交易合計，
    會員等級看 POS_VIP_MAS.CLASS_ID（02 白金、03 鑽石；台灣會員 org_id 是空的）。
    """
    y, m = int(target_month[:4]), int(target_month[4:])
    end = f'{y}-{m:02d}-01'
    start = f'{y - 1}-{m:02d}-01'
    lm_y, lm_m = (y, m - 1) if m > 1 else (y - 1, 12)
    lm_start = f'{lm_y}-{lm_m:02d}-01'
    base, base_lm = F.base_where(start, end, locs), F.base_where(lm_start, end, locs)
    iphone60 = "l.cat4_id=4004 and l.stk_qty<>0 and abs(l.line_total_aftdisc/l.stk_qty)>=60000"

    result = {}

    def merge(df, keys):
        for _, r in df.iterrows():
            row = result.setdefault(F.LOC2S[r['LOC_ID']], {})
            for key in keys:
                value = r[key.upper()]
                row[key] = float(value) if str(value) not in ('nan', 'None') else 0.0

    merge(F.run_epb(f"""select l.loc_id,
        sum(case when l.cat4_id=4004 then l.stk_qty else 0 end) iph,
        sum(case when l.cat4_id=4004 and upper(l.name) like '%PRO MAX%' then l.stk_qty else 0 end) promax,
        sum(case when {iphone60} then l.stk_qty else 0 end) iph60
        from enqsales_bi l where {base} group by l.loc_id"""), ['iph', 'promax', 'iph60'])
    merge(F.run_epb(f"""select l.loc_id, sum(case when {iphone60} then l.stk_qty else 0 end) iph60_lm
        from enqsales_bi l where {base_lm} group by l.loc_id"""), ['iph60_lm'])
    merge(F.run_epb(f"""select loc_id, sum(case when tot>=100000 then 1 else 0 end) n100,
        sum(case when tot>=70000 then 1 else 0 end) n70
        from (select l.loc_id, l.doc_id, sum(l.line_total_aftdisc) tot from enqsales_bi l
              where {base} group by l.loc_id, l.doc_id) group by loc_id"""), ['n100', 'n70'])
    merge(F.run_epb(f"""select l.loc_id,
        round(sum(case when v.class_id='03' then l.line_total_aftdisc else 0 end)) diamond_amt,
        round(sum(case when v.class_id='02' then l.line_total_aftdisc else 0 end)) platinum_amt
        from enqsales_bi l join pos_vip_mas v on v.vip_id=l.cust_id and v.org_id is null
        where {base} group by l.loc_id"""), ['diamond_amt', 'platinum_amt'])
    return result


def push_high_end(target_month, data, ids):
    rows = [{'store_id': ids[store], **values} for store, values in data.items()]
    body = json.dumps({'ym': target_month, 'rows': rows,
                       'pushed_by': os.environ.get('USER', '')}).encode('utf-8')
    request = urllib.request.Request(
        f'{HUB_URL}/api/epb/high-end', data=body, method='POST',
        headers={'Content-Type': 'application/json', 'Authorization': f'Bearer {token()}'})
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            print(f'  ✅ 高單價指標 {target_month}：{json.load(response)}')
    except urllib.error.HTTPError as error:
        sys.exit(f'  ❌ 高單價指標推送失敗 HTTP {error.code}：{error.read().decode("utf-8", "replace")}')
    except urllib.error.URLError as error:
        sys.exit(f'  ❌ 高單價指標 連不到 {HUB_URL}：{error.reason}')


def push(ym, rows):
    body = json.dumps({'ym': ym, 'rows': rows,
                       'pushed_by': os.environ.get('USER', '')}).encode('utf-8')
    request = urllib.request.Request(
        f'{HUB_URL}/api/epb/monthly', data=body, method='POST',
        headers={'Content-Type': 'application/json',
                 'Authorization': f'Bearer {token()}'})
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            print(f'  ✅ {ym}：{json.load(response)}')
    except urllib.error.HTTPError as error:
        sys.exit(f'  ❌ {ym} 推送失敗 HTTP {error.code}：{error.read().decode("utf-8", "replace")}')
    except urllib.error.URLError as error:
        sys.exit(f'  ❌ {ym} 連不到 {HUB_URL}：{error.reason}')


def use_lib_bridge():
    """launchd 讀不到 ~/Desktop（TCC），改用 ~/工具中心/lib 那份 EPBReportQuery（第一個參數是 EPB 根目錄）。"""
    import io as _io
    import subprocess
    import pandas as pd
    java = '/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java'
    lib = os.path.expanduser('~/工具中心/lib')
    classpath = f'{lib}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*'

    def run_epb(sql):
        result = subprocess.run([java, '-cp', classpath, 'EPBReportQuery', '/Library/EPBrowser', sql, '99999'],
                                capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=lib)
        if not result.stdout.strip():
            raise RuntimeError('EPB 沒有回資料（沒連 VPN？）' + result.stderr[-300:])
        return pd.read_csv(_io.StringIO(result.stdout), sep='\t', dtype=str)
    F.run_epb = run_epb


def hub_request(path, body=None):
    """帶 token 打 sa-north-hub 的 API；body 給 dict 就 POST JSON。回傳 (HTTP 狀態, 內容 bytes)。"""
    data = None if body is None else json.dumps(body, ensure_ascii=False).encode('utf-8')
    request = urllib.request.Request(f'{HUB_URL}{path}', data=data, method='GET' if data is None else 'POST',
                                     headers={'Content-Type': 'application/json',
                                              'Authorization': f'Bearer {token()}'})
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return response.status, response.read()
    except urllib.error.HTTPError as error:
        return error.code, error.read()


def intranet_targets(month):
    """內網目前這個月所有店的目標：{店: {kpi代碼: (kpi名稱, 值)}}。"""
    url = f'{INTRANET_URL}?action=query&data_date={month}&store_id=&store_type=nom'
    with urllib.request.urlopen(url, timeout=30) as response:
        rows = json.load(response)
    result = {}
    for row in rows:
        result.setdefault(row['store_id'], {})[row['target_type']] = (row['target_name'],
                                                                     float(row['target_value'] or 0))
    return result


def intranet_upload(file_bytes, filename):
    """跟內網頁面的上傳表單送一樣的東西：file＋action=upload＋store_type=nom。回傳內網的回應文字。"""
    boundary = '----StudioATargetUpload'
    parts = [f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{value}\r\n'.encode()
             for name, value in (('action', 'upload'), ('store_type', 'nom'))]
    parts += [f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="{filename}"\r\n'
              'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n'.encode(),
              file_bytes, f'\r\n--{boundary}--\r\n'.encode()]
    request = urllib.request.Request(INTRANET_URL, data=b''.join(parts), method='POST',
                                     headers={'Content-Type': f'multipart/form-data; boundary={boundary}'})
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read().decode('utf-8', errors='replace')


def check_upload(file_bytes, before, after):
    """回查內網：上傳檔的每一格都要在內網上、值一樣；檔案以外的店要跟上傳前一模一樣。"""
    import openpyxl
    sheet = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True).active
    expected = {(str(row[1]), str(row[2])): (str(row[3]), float(row[4] or 0))
                for row in sheet.iter_rows(min_row=2, values_only=True) if row[1] and row[2]}
    mine = {store for store, _ in expected}
    mismatches = []
    for (store, kpi), (name, value) in expected.items():
        actual = after.get(store, {}).get(kpi)
        if actual is None or abs(actual[1] - value) > 0.005:
            mismatches.append({'store': store, 'kpi': kpi, 'name': name, 'expected': value,
                               'actual': None if actual is None else actual[1]})
    others = sorted((set(before) | set(after)) - mine)
    others_changed = []
    for store in others:
        old, new = before.get(store, {}), after.get(store, {})
        for kpi in sorted(set(old) | set(new), key=lambda k: (len(k), k)):
            if (old.get(kpi) or (None, None))[1] != (new.get(kpi) or (None, None))[1]:
                others_changed.append({'store': store, 'kpi': kpi, 'name': (new.get(kpi) or old.get(kpi))[0],
                                       'before': (old.get(kpi) or (None, None))[1],
                                       'after': (new.get(kpi) or (None, None))[1]})
    return {'expected_cells': len(expected), 'mismatches': mismatches,
            'others_checked': len(others), 'others_changed': others_changed,
            'region_values': {store: {kpi: value for kpi, (_, value) in after.get(store, {}).items()}
                              for store in mine}}


def upload_jobs():
    """目標訂定頁按了「上傳到內網」：上傳 → 回查內網逐格核對、別店有沒有被動到 → 結果回報給頁面。"""
    status, body = hub_request('/api/intranet-upload/pending')
    if status != 200:
        print(f'[{datetime.now():%Y-%m-%d %H:%M}] ⚠️ 讀不到待上傳清單 HTTP {status}', flush=True)
        return
    jobs = json.loads(body).get('pending') or []
    ready = standby_ready('upload', [str(job['id']) for job in jobs])
    for job in jobs:
        if str(job['id']) not in ready:
            continue
        month, label = job['ym'].replace('-', ''), f"{job['ym']} {job['region']}"
        stamp = f'{datetime.now():%Y-%m-%d %H:%M}'
        try:
            before = intranet_targets(month)     # 先確認連得到內網；沒連 VPN 就先不領，下一輪再試
        except Exception as error:
            print(f'[{stamp}] ⚠️ 上傳內網 {label} 先不傳：連不到內網（沒連 VPN？）{str(error)[:120]}', flush=True)
            continue
        status, file_bytes = hub_request(f"/api/intranet-upload/{job['id']}/claim", {})
        if status != 200:
            print(f'[{stamp}] ⚠️ 上傳內網 {label} 領不到工作 HTTP {status}', flush=True)
            continue
        result = {}
        try:
            text = intranet_upload(file_bytes, f'target_nom_{month}.xlsx')
            result['response'] = text[:300]
            try:
                result['response_ok'] = json.loads(text).get('status') == 'success'
            except (ValueError, AttributeError):
                result['response_ok'] = False
        except Exception as error:
            result['error'] = f'上傳時出錯：{str(error)[:200]}'
        try:                                      # 上傳出錯也回查一次，讓頁面看得到內網現在的樣子
            result.update(check_upload(file_bytes, before, intranet_targets(month)))
        except Exception as error:
            result.setdefault('error', f'回查內網時出錯：{str(error)[:200]}')
        ok = bool(result.get('response_ok') and not result.get('error') and result.get('expected_cells')
                  and not result.get('mismatches') and not result.get('others_changed'))
        status, _ = hub_request(f"/api/intranet-upload/{job['id']}/result", {'ok': ok, 'result': result})
        print(f"[{stamp}] {'✅' if ok else '❌'} 上傳內網 {label}：{result.get('error') or ''}"
              f"{result.get('expected_cells', 0) - len(result.get('mismatches') or [])}/{result.get('expected_cells', 0)} 格一致，"
              f"別店變動 {len(result.get('others_changed') or [])} 格（回報 HTTP {status}）", flush=True)


# 心跳要測的兩個內網服務：EPB（Java 橋接打的 SOAP）與內網目標網站。只做 TCP 連線，不到一秒
EPB_HOST, INTRANET_HOST = ('192.168.1.177', 8080), ('192.168.1.186', 80)
DAILY_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '日報_推雲端.py')
CATCHUP_STATE = os.path.expanduser('~/.config/studioa/daily_catchup.json')
CATCHUP_RETRY = 15 * 60          # 同一天補推失敗，15 分鐘後才再試（一次要查 EPB 一分多鐘）
KPI_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '店長KPI_推雲端.py')
KPI_STATE = os.path.expanduser('~/.config/studioa/manager_kpi_retry.json')


def host_name():
    import socket
    return socket.gethostname().split('.')[0][:60]


def reachable(address, timeout=1.5):
    import socket
    try:
        with socket.create_connection(address, timeout=timeout):
            return True
    except OSError:
        return False


def heartbeat():
    """跟雲端報到（同步狀態頁的燈號），回傳 (EPB 連得到嗎, 雲端說缺哪天的日報, 日報頁按了「手動更新」的日期,
    店長 KPI 要（重）推的月份)。"""
    checks = {'epb': reachable(EPB_HOST), 'intranet': reachable(INTRANET_HOST)}
    status, body = hub_request('/api/agent/heartbeat', {'host': host_name(), 'checks': checks})
    reply = json.loads(body) if status == 200 else {}
    return (checks['epb'], reply.get('daily_missing'), reply.get('daily_refresh') or [],
            reply.get('manager_kpi') or [])


def catch_up_daily(missing):
    """雲端缺某天日報（08:30 那次沒連 VPN、Mac 沒開…）→ 這裡補推；同一天失敗 15 分鐘內不重試。"""
    import subprocess
    import time
    try:
        with open(CATCHUP_STATE, encoding='utf-8') as handle:
            state = json.load(handle)
    except (OSError, ValueError):
        state = {}
    if state.get('date') == missing and time.time() - state.get('at', 0) < CATCHUP_RETRY:
        return
    with open(CATCHUP_STATE, 'w', encoding='utf-8') as handle:
        json.dump({'date': missing, 'at': time.time()}, handle)
    stamp = f'{datetime.now():%Y-%m-%d %H:%M}'
    print(f'[{stamp}] 補推門市日報 {missing}', flush=True)
    result = subprocess.run([sys.executable, DAILY_SCRIPT, missing], capture_output=True, text=True,
                            timeout=20 * 60, env={**os.environ, 'LAUNCHD': '1'})
    print((result.stdout + result.stderr).strip()[-500:], flush=True)


def refresh_daily(day):
    """日報頁按了「手動更新」：重查重推那一天；推成功雲端就清旗標，失敗旗標留著、15 分鐘後雲端作廢。"""
    import subprocess
    print(f'[{datetime.now():%Y-%m-%d %H:%M}] 手動更新門市日報 {day}', flush=True)
    result = subprocess.run([sys.executable, DAILY_SCRIPT, day], capture_output=True, text=True,
                            timeout=20 * 60, env={**os.environ, 'LAUNCHD': '1'})
    print((result.stdout + result.stderr).strip()[-500:], flush=True)


def push_manager_kpi(ym):
    """店長 KPI：雲端說這個月要（重）推（每月 1 日自動、頁面按刷新、交辦條件改了）。
    同一個月失敗後 15 分鐘內不重試（一次要查 EPB 半分鐘以上，雲端每輪心跳都會再要一次）。"""
    import subprocess
    import time
    try:
        with open(KPI_STATE, encoding='utf-8') as handle:
            state = json.load(handle)
    except (OSError, ValueError):
        state = {}
    if time.time() - state.get(ym, 0) < CATCHUP_RETRY:
        return
    state[ym] = time.time()
    with open(KPI_STATE, 'w', encoding='utf-8') as handle:
        json.dump(state, handle)
    print(f'[{datetime.now():%Y-%m-%d %H:%M}] 店長 KPI {ym}', flush=True)
    result = subprocess.run([sys.executable, KPI_SCRIPT, ym], capture_output=True, text=True,
                            timeout=20 * 60, env={**os.environ, 'LAUNCHD': '1'})
    print((result.stdout + result.stderr).strip()[-500:], flush=True)
    if result.returncode != 0:
        hub_request('/api/agent/report', {'job': f'店長 KPI {ym[:4]}/{ym[4:]}', 'ok': False, 'host': host_name(),
                                          'message': (result.stdout + result.stderr).strip()[-300:]})


def poll():
    """launchd 每 30 秒跑一次：先報到，雲端有人按「上傳到內網」或「刷新 EPB」才動作；日報缺天就補推。"""
    epb_ok, daily_missing, daily_refresh, kpi_due = False, None, [], []
    try:
        epb_ok, daily_missing, daily_refresh, kpi_due = heartbeat()
    except Exception as error:
        print(f'[{datetime.now():%Y-%m-%d %H:%M}] ⚠️ 心跳沒送出：{str(error)[:200]}', flush=True)
    try:
        upload_jobs()                             # 先做：人在頁面上等結果，而且比重查 EPB 快很多
    except Exception as error:
        print(f'[{datetime.now():%Y-%m-%d %H:%M}] ⚠️ 上傳內網這輪失敗：{str(error)[:200]}', flush=True)
    request = urllib.request.Request(f'{HUB_URL}/api/epb/refresh-pending',
                                     headers={'Authorization': f'Bearer {token()}'})
    with urllib.request.urlopen(request, timeout=30) as response:
        pending = standby_ready('epb', json.load(response).get('pending') or [])
    daily_refresh = standby_ready('daily_refresh', daily_refresh)
    if epb_ok:
        for day in daily_refresh:
            try:
                refresh_daily(day)
            except Exception as error:
                print(f'[{datetime.now():%Y-%m-%d %H:%M}] ⚠️ 手動更新日報出錯：{str(error)[:200]}', flush=True)
    daily_missing = (standby_ready('daily', [daily_missing] if daily_missing else []) or [None])[0]
    if daily_missing and epb_ok and daily_missing not in daily_refresh:
        try:
            catch_up_daily(daily_missing)
        except Exception as error:
            print(f'[{datetime.now():%Y-%m-%d %H:%M}] ⚠️ 補推日報出錯：{str(error)[:200]}', flush=True)
    kpi_due = standby_ready('manager_kpi', kpi_due)
    if epb_ok:
        for ym in kpi_due:
            try:
                push_manager_kpi(ym)
            except Exception as error:
                print(f'[{datetime.now():%Y-%m-%d %H:%M}] ⚠️ 店長 KPI 出錯：{str(error)[:200]}', flush=True)
    if not pending:
        return
    use_lib_bridge()
    for target_month in pending:
        stamp = f'{__import__("datetime").datetime.now():%Y-%m-%d %H:%M}'
        print(f'[{stamp}] 刷新 {target_month}', flush=True)
        try:
            run(target_month)
        except (Exception, SystemExit) as error:
            # 沒連 VPN 時每 2 分鐘都會失敗：只記一行，旗標留著，下次再試
            print(f'[{stamp}] ⚠️ {target_month} 沒推成：{str(error).splitlines()[0][:200]}', flush=True)


def main():
    if len(sys.argv) < 2:
        sys.exit('用法：python3 目標分配_推雲端.py YYYYMM　或　--poll（launchd 用）')
    if sys.argv[1] == '--poll':
        poll()
        return
    run(sys.argv[1].strip())


def run(target_month):
    ly_range, lm_range, ly_label, lm_label = F.month_ranges(target_month)
    # month_ranges 回的標籤是 '2025/10' 這種，補零成 YYYYMM
    ly_ym = '{}{:02d}'.format(*map(int, ly_label.split('/')))
    lm_ym = '{}{:02d}'.format(*map(int, lm_label.split('/')))
    print(f'目標月份 {target_month}：去年同期={ly_label}　上月={lm_label}')

    ids = store_ids()
    locs = F.NORTH1 + F.NORTH2          # 北一北二一起查，雲端兩區共用同一張 epb_monthly
    for ym, date_range, label in ((ly_ym, ly_range, ly_label), (lm_ym, lm_range, lm_label)):
        print(f'查 EPB {label} …')
        data = F.query(*date_range, locs)
        sacare, unmatched = F.query_sacare(*date_range, F.base_where(*date_range, locs))
        F.report_unmatched(label, unmatched)
        rows = rows_for(data, sacare, ids)
        missing = [name for name in ids if name not in data]
        if missing:
            print(f'  ⚠️ {label} 這幾店 EPB 沒有資料，不會推上去：{"、".join(missing)}')
        push(ym, rows)

    print('查 EPB 高單價客群指標（前 12 個月）…')
    push_high_end(target_month, query_high_end(target_month, locs), ids)


if __name__ == '__main__':
    main()
