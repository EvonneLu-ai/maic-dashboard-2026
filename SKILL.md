# MAIC Dashboard Skill

## 基本資訊

| 項目 | 內容 |
|---|---|
| 主要檔案 | `/Users/evonnelu/Desktop/maic-dashboard-deploy/index.html` |
| 分析中心 | `/Users/evonnelu/Desktop/maic-dashboard-deploy/analysis.html` |
| 公開網址 | https://evonnelu-ai.github.io/maic-dashboard-2026/ |
| 分析中心網址 | https://evonnelu-ai.github.io/maic-dashboard-2026/analysis.html |
| GitHub Repo | https://github.com/EvonneLu-ai/maic-dashboard-2026 |

### 部署指令
```bash
cd ~/Desktop/maic-dashboard-deploy
git add -A
git commit -m "Update data YYYY/MM/DD: 說明更新內容"
git push origin main
```
> Push 後約 1～2 分鐘網站自動更新。

---

## 一、標準同步流程（最常用）

### 有新隊伍時
```
「幫我同步 index，帳號 XXX 位、學生 XXX 人」+ 附 CSV（全量）
```

AI 自動執行：
1. 從 CSV 解析所有隊伍（**任一成員統計邏輯**）
2. 重建 analysis.html TEAMS 陣列（含 `extraSchool` 欄位）
3. 計算：總隊數、賽道分佈、schoolChart Top9（B1邏輯）、deptChart、涵蓋學校數
4. 更新 index.html 所有 KPI + 圖表 + 日期
5. 更新 analysis.html 時程/漏斗 TL_DATA、頁首、footer 日期
6. git push 部署
7. 輸出快報摘要

### 只有帳號/學生人數變動時
```
「幫我同步 index，帳號 XXX 位、學生 XXX 人」（不附 CSV）
```
只更新帳號、學生人數、日期。

---

## 二、CSV 格式（全量匯出）

系統匯出的原始 CSV，欄位說明：

| 欄位索引 | Header 名稱 | 實際內容 |
|---|---|---|
| 0 | 組別代號 | 隊伍 ID（如 016） |
| 1 | 組名 | 隊伍名稱 |
| 2 | 賽道 | 創意/創新/創業（隊伍列）或 隊長/隊員/指導老師（成員列）|
| 8 | 電話 | **實際是學校名稱**（系統欄位偏移）|

> ⚠️ Header「電話」欄實際儲存的是「學校 - 系所」，解析時要用 index 8。

---

## 三、學校統計邏輯（方案 B1）

**「任一成員（含指導老師）含該校 = 計入該學校」**

- TEAMS `school` 欄位 = **隊長學校**（顯示用）
- 跨校隊伍有 `extraSchool` 欄位，記錄其他成員的學校
- 學校分佈圖表（index.html + analysis.html）**都計入 extraSchool**
- 兩頁數字完全一致

### 目前（2026/05/19）關鍵數字

| 學校 | 隊數 |
|---|---|
| 臺灣大學 | 61（034 夜行式隊員楊承燁） |
| 逢甲大學 | 49 |
| 政治大學 | 41（含 4 隊政大指導老師 + 0545烤五花肉飯） |
| 臺中教育大學 | 23 |
| 臺北教育大學 | 21 |
| 高雄科技大學 | 21 |

### 跨校隊伍列表（8 隊，有 extraSchool）

| ID | 隊名 | 主學校 | extraSchool |
|---|---|---|---|
| 016 | Innovative!! | 臺北醫學大學 | 國立政治大學 |
| 034 | 夜行式 | 國立清華大學 | 國立臺灣大學, 國立臺中教育大學 |
| 056 | 狐狸 | 逢甲大學 | 國立政治大學 |
| 089 | 大展鴻圖 | 逢甲大學 | 國立政治大學 |
| 0138 | 貓咪 | 國立臺中教育大學 | 國立政治大學 |
| 0299 | 信箱 | 逢甲大學 | 靜宜大學 |
| 0405 | 6年有始有終 | 逢甲大學 | 國立中山大學 |
| 0439 | 窩四立本人 | 國立臺灣科技大學 | 國立臺中科技大學 |

---

## 四、Dashboard（index.html）分頁架構

| 分頁 | 內容 |
|---|---|
| Operation | 報名數據 KPI、科系分佈、學校分佈、活動類型分佈、全年活動排程 |
| Marketing | 宣講會場次、Email 宣傳效益、其他宣傳管道 |
| Training | RTC 課程列表（自動標記已完成）、培訓出席統計、多元能力強化講座 |
| 決議追蹤 | 週會決議事項，含完成／進行中／延誤狀態 |

### KPI 欄位對應

| 要改的項目 | 搜尋關鍵字 |
|---|---|
| 資料截至日期 | `dataAsOf` / `currentDate` |
| 累計報名隊伍 | `累計報名隊伍` |
| 目標達成率 % | `目標達成率` |
| 累計註冊帳號 | `累計註冊帳號` |
| 報名學生人數 | `報名學生人數` |
| 涵蓋學校數 | `涵蓋學校數` |
| 創意/創新/創業組數 | `groupBreakdown` |
| 科系分佈圖 | `deptChart` |
| 學校分佈圖 | `schoolChart` |

> ⚠️ **重要**：index.html 已移除 Google Sheet 對 `accounts`、`students`、`schools` 三個 KPI 的覆蓋邏輯。
> 所有 KPI 直接寫在 JS 內建值，**不需要同時更新 Google Sheet**。
> Google Sheet 僅保留「宣講會」、「決議追蹤」兩個分頁的同步。

### 摘要訊息格式（LINE / Slack 風格）

同步完成後自動輸出：

```
📊 *MAIC 2026 報名快報*
🗓 資料截至 YYYY/MM/DD

━━━━━━━━━━━━━━
🏆 累計報名隊伍：*XXX 組*
📈 目標達成率：*XX.X%*（XXX / 800）

🎯 賽道分佈
　　創意組：XXX 隊
　　創新組：XXX 隊
　　創業組：XXX 隊

🏫 涵蓋學校：*XX 所*
👤 累計帳號：XXX 位
👥 報名人數：XXX 人

━━━━━━━━━━━━━━
🏅 Top 3 學校
　　1. 臺灣大學｜XX 隊（XX%）
　　2. 逢甲大學｜XX 隊（XX%）
　　3. 政治大學｜XX 隊（XX%）

━━━━━━━━━━━━━━
🔗 完整 Dashboard：
   https://evonnelu-ai.github.io/maic-dashboard-2026/
━━━━━━━━━━━━━━━━━━━━
Apple EDU Taiwan · Evonne Lu
```

---

## 五、其他常見更新任務

**① 新增宣講會（Marketing Tab）**

在 JS 中找 `sessions:` 陣列，新增一筆：
```javascript
{ school:'學校名稱', date:'日期', lecturer:'講師', type:'線下/線上',
  time:'時間', people:人數, venue:'地點', note:'辦理情況', color: C.顏色 }
```

**② 新增 RTC 課程（Training Tab）**

在 JS 中找 `rtc:` 陣列，新增一筆：
```javascript
{ date:'M/D(週)', name:'課程名稱', lecturer:'講師', done:false, people:0 }
```
> ⚠️ `done` 欄位由 `autoMarkTrainingDone()` 自動依日期標記，手動設定僅供強制覆蓋。
> `people` 欄位填入實際出席人數，會自動加總到「培訓累計出席」KPI。

**③ 新增多元講座（Training Tab）**

在 JS 中找 `workshops` 陣列，在對應月份的 `items` 中新增：
```javascript
{ date:'M/D(週)', name:'講座名稱', format:'實體/線上',
  lecturer:'講師名稱', people:人數, meta:'課程描述（選填）', done:false }
```
> `lecturer`、`people`、`meta` 欄位為選填，有值才會在卡片上顯示。
> `people` 會自動加總到「多元講座出席」KPI。

**④ 新增決議事項（決議追蹤 Tab）**

在 JS 中找 `decisions` 陣列，在對應週次的 `items` 中新增：
```javascript
{ id:'dXX', owner:'負責人', item:'決議內容', deadline:'M/DD', priority:'high/mid/low', status:'open' }
```

---

## 六、分析中心（analysis.html）TEAMS 陣列

### 欄位格式
```javascript
{id:'16', name:'Innovative!!', track:'創業', school:'臺北醫學大學',
 region:'北部', dept:'醫學系', deptCat:'法律醫療', nameCat:'英文混搭',
 extraSchool:'國立政治大學'}  // 只有跨校隊伍才有此欄位
```

### 欄位值對照

| 欄位 | 可能的值 |
|---|---|
| `track` | `創意` / `創新` / `創業` |
| `region` | `北部` / `中部` / `南部` / `其他` |
| `deptCat` | `資工電機` / `商管` / `教育` / `設計藝術` / `語文人文` / `理工科學` / `法律醫療` / `其他` |
| `nameCat` | `科技梗` / `食物飲料` / `動物` / `幽默自嘲` / `英文混搭` / `自然意象` / `日常生活` / `戰鬥冒險` / `其他` |

### `countBySchool()` 函式（方案 B1 核心）
```javascript
function countBySchool(arr) {
  const r={};
  arr.forEach(t=>{
    if(t.school) r[t.school]=(r[t.school]||0)+1;
    if(t.extraSchool) t.extraSchool.split(',').forEach(s=>{
      s=s.trim(); if(s) r[s]=(r[s]||0)+1;
    });
  });
  return r;
}
```
> 所有 `countBy(arr, 'school')` 已替換為 `countBySchool(arr)`

### 時程/漏斗（TL_DATA）
- 位於 analysis.html 底部 `const TL_DATA = [...]`
- 每筆格式：`{date:'區間', count:隊數, event:'觸發事件', color:'#hex'}`
- 佔比分母 = 總隊伍數（每次更新需同步改 `t.count/415*100` 中的分母）
- 漏斗 `已完成基本報名` count 也要改為最新總隊數

---

## 七、關鍵數據（2026/05/19）

| 項目 | 數值 |
|---|---|
| 總隊伍 | 415 |
| 創意組 | 195 |
| 創新組 | 177 |
| 創業組 | 43 |
| 涵蓋學校數 | 51 所 |
| 帳號 | 505 位 |
| 學生人數 | 505 人 |
| 目標達成率 | 51.9%（415/800）|
| 0260 聲映初晨 | 完全無學校資料（創業組，僅有報名列）|

### Marketing 培訓概況（截至 2026/05/18）

| 項目 | 數值 |
|---|---|
| 宣講會暨工作坊場次 | 12 場（線下 9 + 線上 3）|
| 宣講累計出席 | 364 人 |
| 最新場次 | 5/15 逢甲大學工作坊 |

### Training 培訓概況

| 項目 | 數值 |
|---|---|
| RTC 課程總場次 | 10 場 |
| RTC 已完成 | 7 / 10（截至 5/15）|
| 多元講座出席（已登錄）| 110 人（3/14×70 + 5/13×30 + 5/15×10）|
| RTC 課程出席（待補）| 各場人數尚未登錄，`people:0` 待更新 |

### 歷史傳承種子
| 姓名 | 隊伍 | 學校 | 年級 |
|---|---|---|---|
| 鄭凱馨 | 034「夜行式」隊長 | 清大 教育與學習科技學系 | 碩士生 |
| 楊承燁 | 034「夜行式」隊員 | 台大 電機工程學系 | 大三 |
| 黃廷宇 | 037「每個人都想成為我的豎笛」隊長 | 逢甲 資訊工程學系 | 大四 |

---

## 八、Training Tab 架構（2026/05/19 更新）

### KPI 行列
- **第一列（k2）**：RTC 課程總場次 / 已完成場次
- **第二列（k3）**：RTC 課程出席 / 多元講座出席 / 培訓累計出席
  - 第二列為**自動計算**，加總 `rtc[].people` + `workshops[].items[].people`

### 多元講座卡片顯示欄位
卡片會顯示（有值才出現）：
- 日期、課程名稱、✅/⏳ 狀態
- 👤 `lecturer` 講師名稱
- 👥 `people` 人數
- `meta` 課程描述（分隔線下方）
- 形式 badge（實體/線上）

---

## 九、顏色對照

```
C.blue   = #0071e3   ← Operation
C.orange = #ff9500   ← Marketing
C.green  = #34c759   ← Training / 已完成
C.red    = #ff3b30   ← 延誤 / 警示
C.purple = #af52de   ← 決議追蹤
C.teal   = #5ac8fa   ← 輔助色
C.pink   = #ff2d55   ← 輔助色
```
