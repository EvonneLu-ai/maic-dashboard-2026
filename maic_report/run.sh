#!/bin/bash
# ╔══════════════════════════════════════════════╗
# ║  MAIC 2026 週報一鍵產出                       ║
# ║  用法：bash run.sh 本週.csv [上週.csv]         ║
# ╚══════════════════════════════════════════════╝

cd "$(dirname "$0")"

if [ -z "$1" ]; then
  echo ""
  echo "❌ 請指定本週 CSV 檔案名稱"
  echo ""
  echo "  用法：bash run.sh 本週.csv"
  echo "  進階：bash run.sh week_0428.csv week_0421.csv"
  echo ""
  exit 1
fi

python3 generate_report.py "$@"

echo ""
echo "📂 開啟輸出資料夾..."
open output/
