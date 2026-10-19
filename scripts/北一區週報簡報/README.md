# 北一區／北二區週報簡報自動產生

兩區共用同一套程式，`--region 北二區` 切換（設定在 build_deck.py 的 `REGIONS`）。北二區差異：7 店、週次寫 Q1W3、門市週報 Excel 沒有自動寄信（與北一區 dss@ 信件附件同一個工具產生）→改找下載資料夾的 `北二區_北二區週報_YYYYMMDD.xlsx`、排班系統看不到→銷售人數用本週有銷售的人數（黃底）、3PP 主表是公式→寫進頁外來源表「表格 1-1」（依門市代碼順序、比率存 300）、維修只有花蓮 113（歷史月份由 EPB 重算）、簡報在 `桌面/週報/北二區/`。

每週一 12:30 由 Claude 桌面 App 排程任務「北一區週報簡報」執行，13:00 前交件：
複製桌面/週報 裡上一週的 `北一區週報 XXQ#W##.key` → 填入本週數字、換圖、改週次 → 寫評語 → 匯出 PDF。

```bash
python3 build_deck.py                       # 數字＋圖表（上週日~週六）
python3 fill_comments.py 27Q1W03            # 讀 評語.json 寫進簡報並匯出 PDF
python3 fill_comments.py 27Q1W03 --pdf-only # 手改簡報後只重新匯出 PDF
```
`build_deck.py --wk-end 2026-10-10 --out-dir /tmp/x` 可指定週、輸出到別處測試。

## 檔案
- `build_deck.py` 撈資料、組 Keynote 修改指令、產生 `評語素材.md`
- `kn.js` Keynote 橋接（JXA）：`dump` 讀結構、`apply` 依 ops.json 修改＋存檔＋匯出 PDF
- `chart_gui.py` 維修中心長條圖：保留原生圖表（樣式與原版完全相同），模擬人工操作「圖表資料編輯器」
  貼上新資料，再全選拷貝回來逐格核對；分隔虛線與「去年／今年」字樣由 kn.js `yearmark` 跟著柱數移位。
  執行時會操作 Keynote 畫面（每張約 25 秒），螢幕不可鎖定、不要動鍵盤滑鼠；剪貼簿暫用後還原文字。
- `charts.py` 備援：chart_gui 任何一張失敗就整份放棄畫面修改，改用圖片取代圖表（樣式略有不同，會列入 warnings）；
  圖片 description 標 `AUTO_CHART_<3pp|repair>_<代碼>`，之後的週次依此找到並替換
- `fill_comments.py` 寫評語、匯出 PDF
- `finals_mail.py` 北二區：從信箱取回 Yuni 每週寄的「北二區 27Q1W# 週報」定稿 .key（同週多封取最後一封），
  存到 `輸出/北二區週報簡報/定稿信件/`；`find_template` 同一週優先用它當範本，評語學習也只拿它當定稿
  （還沒收到時用自己上週的簡報當範本、跳過學習並列 warning）
- `評語寫作規則_北二區.md` Yuni 寫法摘要（由 39 週定稿歸納）
- `評語寫作規則.md`、`評語範例/` 評語寫法（從使用者 26Q4～27Q1 手寫評語歸納）
- `評語修改紀錄.md` 每週自動比對「上週自動初稿 vs 使用者改完的定稿」（定稿＝下週範本簡報裡的評語），
  寫評語前必讀；定稿也會自動存進 `評語範例/`
- `維修圖表歷史.json` 維修長條圖的歷史月份（使用者原圖數字，口徑不一無法重算）；新的完整月份自動補上
- 第 9 頁「重點項目」每月會換：`build_deck.py` 的 `KEY_ITEMS` 有定義的品項自動算，其餘標黃底
- 產出與中間檔：`~/工具中心/輸出/北一區週報簡報/<週次>/`（週報 Excel、ops.json、評語素材.md、評語.json、PDF、state.json）

## 各頁來源
| 頁 | 來源 |
|---|---|
| 業績總表 | dss@ 每週日 04:05 寄的「北一區門市週報_YYYYMMDD」附件，照其『週報』分頁公式重算；來客數＝週報產生器人流；本週銷售人數＝排班系統（本週有排任何格子的人，羅東維修除外），休假固定 0 |
| 重點項目進度 | 週報產生器 Excel「重點項目進度」（AC+ 佔比自己用分子分母算）；Mac mini M4＝EPB 本月 cat3=3001 品名含 MAC MINI M4 |
| MySetup | `warranty_report_all.scan_mysetup`（Personal Setup 信）；主機總數、服務費＝週報 Excel「本週其他細項」 |
| 3PP 搭售統計 | `dss_rebuild_from_epb`；每月第一週自動在 Total 下插上個月整月列 |
| 維修營收表、圖 | `repair_report`（EPB 士林 110、羅東 121） |
| **黃底手填** | 維修中心預約、SACare 整機搭售、CSAT —— 來源是週二才寄的維修週報（Mike 的 ASP 週報、士林維修週報） |

## 已知差異（2026-10-13 驗證 27Q1W02）
- 3PP：EPB 重建對 DSS 匯出 83/90 格相同；差異都是職業 88/99 公司戶「銷售→銷退→訂金→分期交貨」的單，DSS 有算、重建排除。樣本少，不調規則。
- MySetup 士林 46（手做版 47）。

## 注意
- 要開公司 VPN（EPB）；Mail.app 要已收到 dss@ 的信。
- 範本與輸出都在 iCloud 桌面，launchd 讀不到，所以用 Claude 排程任務跑而不是 launchd。
- 排程時間 Mac 要醒著、Claude App 要開著（沒開會在下次開啟時補跑）。
