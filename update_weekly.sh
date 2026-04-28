#!/bin/bash
# ══════════════════════════════════════════════════════
#  MAIC 報名周報 一鍵更新腳本
#  用法：./update_weekly.sh <CSV檔案路徑>
#  範例：./update_weekly.sh ~/Downloads/members.csv
# ══════════════════════════════════════════════════════

set -e  # 任何步驟失敗立即中止

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WEEKLY_SCRIPT="$HOME/Documents/Enchanté/Conversations/EC59CDFD-6847-4623-A764-7245F5716472/maic_weekly.py"
REPORT_SRC="$HOME/Documents/Enchanté/Conversations/EC59CDFD-6847-4623-A764-7245F5716472/maic_full_report.html"
REPORT_DEST="$SCRIPT_DIR/weekly-report.html"
TODAY=$(date +%Y/%m/%d)

# ── 檢查 CSV 參數 ──────────────────────────────────────
if [ -z "$1" ]; then
  echo ""
  echo "❌ 請指定 CSV 檔案路徑"
  echo "   用法：./update_weekly.sh <CSV檔案路徑>"
  echo "   範例：./update_weekly.sh ~/Downloads/members.csv"
  echo ""
  exit 1
fi

CSV_PATH="$1"

if [ ! -f "$CSV_PATH" ]; then
  echo "❌ 找不到 CSV 檔案：$CSV_PATH"
  exit 1
fi

echo ""
echo "════════════════════════════════════════"
echo "  MAIC 報名周報 一鍵更新"
echo "  資料來源：$CSV_PATH"
echo "════════════════════════════════════════"
echo ""

# ── Step 1：產出報表 ────────────────────────────────────
echo "📊 Step 1/3  分析資料、產出報表..."
cd "$(dirname "$WEEKLY_SCRIPT")"
python3 "$WEEKLY_SCRIPT" "$CSV_PATH"
echo ""

# ── Step 2：複製到 deploy 資料夾 ────────────────────────
echo "📁 Step 2/3  複製報表到 deploy 資料夾..."
cp "$REPORT_SRC" "$REPORT_DEST"
echo "   ✅ weekly-report.html 已更新"
echo ""

# ── Step 3：Git push ────────────────────────────────────
echo "🚀 Step 3/3  Push 到 GitHub..."
cd "$SCRIPT_DIR"
git add weekly-report.html
git commit -m "Update 週報 $TODAY"
git push origin main
echo ""

echo "════════════════════════════════════════"
echo "  ✅ 完成！約 1～2 分鐘後網站自動更新"
echo ""
echo "  🌐 https://evonnelu-ai.github.io/maic-dashboard-2026/weekly-report.html"
echo "════════════════════════════════════════"
echo ""
