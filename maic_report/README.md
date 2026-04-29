# MAIC 2026 每週報告自動產出系統

## 使用方式

```bash
cd ~/Desktop/maic-dashboard-deploy/maic_report

# 基本用法（只用本週CSV）
python generate_report.py sample_data.csv

# 進階：帶上週CSV，自動計算週增量
python generate_report.py week_0428.csv week_0421.csv
```

## CSV 欄位格式

| 欄位 | 說明 | 範例 |
|---|---|---|
| 隊伍ID | 唯一識別碼 | T001 |
| 隊伍名稱 | 隊伍/專題名稱 | 智能學習助理 |
| 學校 | 學校全名 | 國立臺灣大學 |
| 科系 | 科系全名 | 資訊工程學系 |
| 賽道 | 創意 / 創新 / 創業 | 創新 |
| 學生人數 | 隊員數（不含老師） | 3 |
| 指導老師數 | 0 或正整數 | 1 |
| 報名日期 | YYYY-MM-DD 格式 | 2026-04-28 |

## 輸出檔案

| 檔案 | 用途 |
|---|---|
| `output/maic_report_YYYYMMDD.html` | 互動報表，上傳後主管自助查閱 |
| `output/maic_pdf_YYYYMMDD.html` | 瀏覽器開啟→列印→儲存PDF，Email附件用 |
| `output/maic_slides_YYYYMMDD.pptx` | 週會簡報，4張核心投影片 |
| `output/maic_summary_YYYYMMDD.txt` | 複製貼到 LINE / Slack / Teams |

## 安裝套件（首次）

```bash
pip install python-pptx
```

> `python-pptx` 只用於 PPTX 輸出，沒安裝會跳過此格式，其他3種仍正常產出。

---

*MAIC 2026 · Apple EDU Taiwan · Evonne Lu*
