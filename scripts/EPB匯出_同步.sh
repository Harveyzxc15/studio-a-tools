#!/bin/bash
# EPBrowser 匯出檔 → Mac mini（launchd com.bob.epb-export-sync 監看資料夾觸發）
# 匯出檔存到 ~/Downloads/EPB匯出（捷徑）＝ ~/工具中心/匯入/EPB匯出；家裡走區網，外面走 Tailscale。
# 同步完由 mini 用它自己的盲測 webhook 推 Discord（筆電不放 webhook）。
DIR="$HOME/工具中心/匯入/EPB匯出"
LOG_DIR="$HOME/工具中心/輸出/EPB匯出同步"
mkdir -p "$LOG_DIR"
exec >>"$LOG_DIR/sync.log" 2>&1
echo "== $(date '+%F %T')"

# 等檔案寫完：xlsx 總大小 10 秒內不變（最多等 10 分鐘）；Downloads 根目錄的「銷售分析作業*.xlsx」一起算
DL="$HOME/Downloads"
prev=""
for _ in $(seq 1 60); do
  cur=$( { find "$DIR" -maxdepth 1 -name '*.xlsx' ! -name '~$*'; find "$DL" -maxdepth 1 -name '銷售分析作業*.xlsx'; } \
        | while read -r f; do stat -f %z "$f"; done 2>/dev/null | awk '{s+=$1} END{print s+0}')
  [ "$cur" = "$prev" ] && break
  prev=$cur
  sleep 10
done

# 存到 Downloads 根目錄的匯出檔（EPBrowser 預設檔名）自動搬進匯出資料夾
# 只搬 2 小時內存的檔：Downloads 裡以前測試匯出的舊檔不要搬（11/18 把 11/17 下午兩份舊檔搬進去了）
find "$DL" -maxdepth 1 -name '銷售分析作業*.xlsx' -mmin -120 | while read -r f; do
  mv -n "$f" "$DIR/"                     # macOS mv -n 撞名時不搬但也回 0，所以看原檔還在不在
  [ -e "$f" ] && echo "沒搬（匯出資料夾已有同名檔）：$(basename "$f")" || echo "搬進匯出資料夾：$(basename "$f")"
done

SSH="ssh -o ConnectTimeout=8 -o BatchMode=yes -o HostKeyAlias=HdeMac-mini.local"
NOTIFY='import json,sys,urllib.request,pathlib
u=pathlib.Path("~/.config/studioa/mail_digest_discord_webhook.txt").expanduser().read_text().strip()
urllib.request.urlopen(urllib.request.Request(u,data=json.dumps({"content":sys.argv[1]}).encode(),headers={"Content-Type":"application/json","User-Agent":"epb-export-sync"}),timeout=20)'

for host in HdeMac-mini.local 100.66.101.55; do
  out=$(rsync -a -8 --timeout=120 --itemize-changes -e "$SSH" --exclude='~$*' --exclude='.*' --include='*.xlsx' --exclude='*' \
        "$DIR/" "h@$host:工具中心/匯入/EPB匯出/") || { echo "rsync $host 失敗"; continue; }
  new=$(echo "$out" | awk '/^<f/ {sub(/^[^ ]+ /,""); print}')
  echo "rsync $host OK：${new:-（沒有新檔）}"
  if [ -n "$new" ]; then
    msg="📥 EPB 匯出檔已同步到 mini：$(echo "$new" | awk 'NR>1{printf "、"} {printf "%s", $0}')"   # paste -d 只吃分隔字的第一個位元組，「、」會變亂碼
    $SSH "h@$host" "/usr/bin/python3 -c '$NOTIFY' \"$msg\"" || echo "Discord 通知失敗"
  fi
  exit 0
done
echo "兩條路都連不到 mini"
exit 1
