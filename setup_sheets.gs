/**
 * MAIC Dashboard — Google Sheets 初始設定腳本
 * ─────────────────────────────────────────────
 * 執行方式：
 *   1. 在 Google Sheets 上方選單 → 擴充功能 → Apps Script
 *   2. 貼上此腳本全文
 *   3. 點選 ▶ 執行「setupMAICTabs」
 *   4. 授權後等待完成（約 5 秒）
 * ─────────────────────────────────────────────
 */

function setupMAICTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _setupSettings(ss);
  _setupSessions(ss);
  _setupDecisions(ss);

  SpreadsheetApp.getUi().alert('✅ 完成！三個 Tab 已建立：設定、宣講會、決議\n\n之後只需修改這些 Tab 的數字，重新整理 Dashboard 即可自動同步。');
}

// ════════════════════════════════════════════
//  Tab 1：設定（key-value 格式）
// ════════════════════════════════════════════
function _setupSettings(ss) {
  let sheet = ss.getSheetByName('設定');
  if (!sheet) sheet = ss.insertSheet('設定');
  sheet.clearContents();

  const rows = [
    // ── 標題列 ──────────────────────────
    ['key', 'value', '說明'],

    // ── 全域設定 ────────────────────────
    ['currentDate',           '2026-04-17',    '今日日期（YYYY-MM-DD，影響課程自動標記）'],
    ['dataAsOf',              '2026/04/17',    '資料截至日期（顯示用）'],
    ['registrationDateLabel', '截至 2026/04/15', '報名截止日期標籤'],
    ['registrationTarget',    '800',           '報名目標隊伍數'],

    // ── Operation KPI ───────────────────
    ['teams',           '89',                          '累計報名隊伍'],
    ['teamsSub',        '本週新增 +84 隊',               '隊伍 KPI 副標題'],
    ['targetRate',      '11.1%',                       '目標達成率'],
    ['targetRateSub',   'TW Season 初期階段（89/800）',  '達成率副標題'],
    ['accounts',        '142',                         '累計註冊帳號'],
    ['accountsSub',     '已驗證帳號 142 位',              '帳號副標題'],
    ['students',        '142',                         '報名學生人數'],
    ['studentsSub',     '參加過 3 人（2.11%）',           '學生副標題'],
    ['schools',         '12',                          '涵蓋學校數'],
    ['schoolsSub',      '全台各地學校',                  '學校副標題'],

    // ── 組別分佈 ────────────────────────
    ['creative',        '44',   '創意組'],
    ['innovative',      '40',   '創新組'],
    ['startup',         '50',   '創業組'],
    ['submitted',       '1',    '企劃書已繳'],
    ['notSubmitted',    '88',   '企劃書未繳'],

    // ── 網站流量 ────────────────────────
    ['website1Value',   '16,437',                        'maic.straighta.com.tw 瀏覽數'],
    ['website1Sub',     '今日 +63 次瀏覽',                '網站1 副標題'],
    ['website2Value',   '9,199',                         'bts.straighta.com.tw 瀏覽數'],
    ['website2Sub',     '今日 +42・附件下載 198/126/124',  '網站2 副標題'],

    // ── Marketing KPI ───────────────────
    ['mk_sessions',     '6',        '已舉辦宣講會場次'],
    ['mk_sessionsSub',  '線下 5 + 線上 1', '場次副標題'],
    ['mk_maxPeople',    '150+',     '最大單場人數'],
    ['mk_maxPeopleSub', '高雄科技大學 3/18', '最大場次地點'],
    ['mk_totalAttend',  '265+',     '宣講累計出席'],
    ['mk_totalSub',     '6 場合計實際出席', '出席副標題'],
    ['mk_nextSchool',   '清華大學',  '下一場宣講學校'],
    ['mk_nextDate',     '4月下旬 TBC', '下一場日期'],

    // ── Training KPI ────────────────────
    ['tr_total',        '10',  'RTC 課程總場次'],
    ['tr_totalSub',     '3月 ~ 6月，每週五', '課程期間'],
    ['tr_done',         '5',   '已完成場次'],
    ['tr_doneSub',      '截至 4/10（含今日課程）', '完成場次副標題'],
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);

  // 格式化
  const header = sheet.getRange(1, 1, 1, 3);
  header.setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 320);
  sheet.setFrozenRows(1);

  // 說明欄底色
  sheet.getRange(2, 3, rows.length - 1, 1).setFontColor('#888888').setFontStyle('italic');
}

// ════════════════════════════════════════════
//  Tab 2：宣講會
// ════════════════════════════════════════════
function _setupSessions(ss) {
  let sheet = ss.getSheetByName('宣講會');
  if (!sheet) sheet = ss.insertSheet('宣講會');
  sheet.clearContents();

  const rows = [
    ['學校', '日期', '講師', '類型', '時間', '人數', '場地', '備註', '顏色'],
    ['高雄科技大學',      '3/18',     'YY',             '線下', '14:00–14:30', 150,  '燕巢校區',              '學生反應熱烈，活動結束前紛紛掃 LINE QR Code',          '#ff9500'],
    ['臺灣大學（工作坊）', '3/19',     'Ethan / Brandi', '線下', '09:30–12:30', 15,   '臺大共同教學館 1F 103', '所有學生認真參與，對賽事高度興趣，活動後前來詢問',       '#0071e3'],
    ['文化大學',          '3/20',     'Edick & Brandi', '線下', '10:00–11:00', 30,   '資訊傳播學系教室',       '分享得獎作品時，學生表現出高度興趣',                   '#af52de'],
    ['海洋大學',          '3/25',     'Brandi',         '線下', '12:20–12:40', 15,   '資訊工程學系',          '分享 Apple 創業基金後，學生興趣大增',                  '#34c759'],
    ['慈濟大學',          '4/1',      'Lynn',           '線下', '16:10–16:30', 45,   '經管系 自媒體傳播課程',  '老師安排於課程中分享，學校及課程另外提供參賽及獲獎獎勵',  '#ff3b30'],
    ['東華大學（線上）',   '4/2+4/9',  'Lynn',           '線上', '12:30–13:30', 10,   '線上',                 '學生敢於分享瞭解賽事的原因，積極了解比賽及參賽方式',     '#5ac8fa'],
  ];

  sheet.getRange(1, 1, rows.length, 9).setValues(rows);

  const header = sheet.getRange(1, 1, 1, 9);
  header.setBackground('#ff9500').setFontColor('#ffffff').setFontWeight('bold');
  [1,2,3,4,5,7,8].forEach(c => sheet.setColumnWidth(c, 140));
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(9, 90);
  sheet.setFrozenRows(1);
}

// ════════════════════════════════════════════
//  Tab 3：決議（攤平格式）
// ════════════════════════════════════════════
function _setupDecisions(ss) {
  let sheet = ss.getSheetByName('決議');
  if (!sheet) sheet = ss.insertSheet('決議');
  sheet.clearContents();

  const rows = [
    ['會議', '負責人', '事項', '截止日', '優先度', '狀態'],

    // W16
    ['週會 W16｜2026/04/17', '正奎',           '發送資料給中教大郭校長與中心，並通知 Edick 追蹤',                      '4/22',       'high', 'open'],
    ['週會 W16｜2026/04/17', '正奎',           '上傳 GC K-12 賽事規則資料至共享資料夾',                                '4/17',       'high', 'open'],
    ['週會 W16｜2026/04/17', '正奎 / SA',      '向老闆報備 K-12 賽道狀況',                                             '4/21',       'high', 'open'],
    ['週會 W16｜2026/04/17', '正奎',           '確認夏季訓練營 SA 具體分工，並回報是否參與台大協調會',                  '4/24',       'mid',  'open'],
    ['週會 W16｜2026/04/17', 'Bella',          '製作 4/30 台大宣講會 IG / FB 圖卡（需含學校官稱）',                    '儘速',       'high', 'open'],
    ['週會 W16｜2026/04/17', 'Edick',          '與清大邱教授初步聯繫，評估小型宣講會可能性',                           '下週',       'high', 'open'],
    ['週會 W16｜2026/04/17', 'Kat / SA',       '將 PDF 行銷計畫書內容填入 Keynote 簡報模組（含歷年數據）',             '4/24',       'high', 'open'],
    ['週會 W16｜2026/04/17', 'Willy / Brendi', '提供主視覺 AI 原始檔給正奎及傳揚（David）',                           '儘速',       'high', 'open'],
    ['週會 W16｜2026/04/17', 'Brendi',         '更新 Numbers 欄位（標註線上／實體課程），並追蹤創業組名單',             '每週一更新',  'mid',  'open'],
    ['週會 W16｜2026/04/17', 'Evonne',         '串接 Dashboard，產出視覺化週報圖表',                                   '每週會後',   'mid',  'open'],

    // W15 (4/10)
    ['週會 W15｜2026/04/10', 'SA 團隊', '補齊資料，產出 PDF 版賽事規劃書並寄給 Max（CC 正奎、Evonne）',        '4/14', 'high', 'open'],
    ['週會 W15｜2026/04/10', 'SA 團隊', '更新受眾名單並發送新一波 EDM（主視覺確認後執行，導入分眾客製化策略）', '4/17', 'mid',  'open'],
    ['週會 W15｜2026/04/10', 'Edick',   '回覆成功大學場地費洽談進展',                                          '4/17', 'mid',  'open'],
    ['週會 W15｜2026/04/10', 'Edick',   '確認逢甲大學 4/24 宣講會講師費用支付對象（SA 或逢甲大學買單）',        '4/17', 'mid',  'open'],
    ['週會 W15｜2026/04/10', 'Brandi',  '在追蹤表格新增「已註冊人數」欄位',                                    '4/10', 'low',  'done'],
    ['週會 W15｜2026/04/10', '正奎',    '牽線聯繫「傳揚」窗口，收集前年台大決賽場地規劃結案資料',              '4/15', 'mid',  'open'],
    ['週會 W15｜2026/04/10', 'Brandi',  '提醒 Edick 下週三（4/15）下午 2:00 中教大拜訪行程',                  '4/14', 'mid',  'done'],
    ['週會 W15｜2026/04/10', 'Brandi',  '向易璇索取臺大宣講會學生名單（易璇過往已提供，下週一更新名單並發送賽事資訊）', '4/21', 'low', 'open'],
    ['週會 W15｜2026/04/10', '正奎',    '協助處理中教大 Logo 向量檔（AI 格式），解決 PNG 解析度不佳問題',      '4/14', 'mid',  'open'],
    ['週會 W15｜2026/04/10', '正奎',    '確認清華大學宣講會執行可能性（對口：邱富源教授）',                    '4/17', 'mid',  'open'],
    ['週會 W15｜2026/04/10', '正奎',    '將 Kat 與 Amy 的信箱加入試算表與簡報共編權限',                       '4/10', 'low',  'done'],

    // W15 (4/9)
    ['週會 W15｜2026/04/09', 'Brandi', '清華大學宣講時間確認，回覆給 Max',             '4/15', 'high', 'open'],
    ['週會 W15｜2026/04/09', 'Lynn',   '東華大學宣講後有興趣同學名單整理，提醒報名',    '4/16', 'high', 'open'],
    ['週會 W15｜2026/04/09', '13',     '成功大學宣講等 Edick 確認後排定時間',           '4/20', 'mid',  'open'],
    ['週會 W15｜2026/04/09', '全體',   '企劃書未繳 4 組隊伍發提醒信',                  '4/18', 'mid',  'open'],
    ['週會 W15｜2026/04/09', 'YY',     '逢甲大學 4/24 宣講場地確認與講義準備',          '4/22', 'mid',  'open'],

    // W14
    ['週會 W14｜2026/04/02', 'Lynn',   '慈濟大學宣講投影片更新（加入 Apple 創業基金資訊）', '4/05', 'mid', 'done'],
    ['週會 W14｜2026/04/02', 'Brandi', '海洋大學宣講後有興趣同學 follow-up',            '4/08', 'low',  'done'],
    ['週會 W14｜2026/04/02', '13',     'RTC W4（3/27）課程投影片上傳至 BTS 平台',       '4/05', 'low',  'done'],

    // W13
    ['週會 W13｜2026/03/26', 'Edick', '文化大學後續接洽，安排第二次工作坊',           '4/15', 'mid',  'delay'],
    ['週會 W13｜2026/03/26', '全體',  '更新 maic.straighta.com.tw 首頁報名截止倒數', '3/31', 'high', 'done'],
  ];

  sheet.getRange(1, 1, rows.length, 6).setValues(rows);

  const header = sheet.getRange(1, 1, 1, 6);
  header.setBackground('#af52de').setFontColor('#ffffff').setFontWeight('bold');

  // 狀態欄著色
  for (let i = 2; i <= rows.length; i++) {
    const cell = sheet.getRange(i, 6);
    const val = rows[i - 1][5];
    if (val === 'done')  cell.setBackground('#d4edda');
    if (val === 'delay') cell.setBackground('#fff3cd');
    if (val === 'open')  cell.setBackground('#ffffff');
  }

  // 優先度欄著色
  for (let i = 2; i <= rows.length; i++) {
    const cell = sheet.getRange(i, 5);
    const val = rows[i - 1][4];
    if (val === 'high') cell.setFontColor('#dc3545').setFontWeight('bold');
    if (val === 'mid')  cell.setFontColor('#fd7e14');
    if (val === 'low')  cell.setFontColor('#6c757d');
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 90);
  sheet.setFrozenRows(1);
}
