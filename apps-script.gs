// ════════════════════════════════════════════════════════
//  Google Apps Script — 打卡紀錄
//  貼到 Apps Script 編輯器，部署為「網頁應用程式」
//  執行身分：我　｜　存取權限：所有人（包含匿名）
// ════════════════════════════════════════════════════════

function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action    || '';

  if (action === 'clockIn')  handleClockIn(p);
  if (action === 'clockOut') handleClockOut(p);

  // 必須回傳文字，否則 Apps Script 會報錯
  return ContentService
    .createTextOutput('ok')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── 上班打卡 ─────────────────────────────────────────────
function handleClockIn(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sh    = getOrCreate(ss, '打卡紀錄');
  const today = p.date || fmtDate(new Date());

  // 確保有標題列
  if (sh.getLastRow() === 0) {
    sh.appendRow(['姓名', '日期', '上班時間', '下班時間']);
  }

  // 若今日已有紀錄則更新，否則新增
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.name && String(data[i][1]) === today) {
      sh.getRange(i + 1, 3).setValue(p.inTime);
      sh.getRange(i + 1, 4).setValue('');
      return;
    }
  }
  sh.appendRow([p.name, today, p.inTime, '']);
}

// ── 下班打卡 ─────────────────────────────────────────────
function handleClockOut(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sh    = getOrCreate(ss, '打卡紀錄');
  const today = p.date || fmtDate(new Date());

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.name && String(data[i][1]) === today) {
      sh.getRange(i + 1, 4).setValue(p.outTime);
      return;
    }
  }
  // 找不到上班紀錄就補一筆
  sh.appendRow([p.name, today, p.inTime || '', p.outTime]);
}

// ── 工具 ─────────────────────────────────────────────────
function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function fmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd');
}
