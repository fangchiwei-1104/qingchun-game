const TEST_VERSION = 'geology-6-v1';
const UNIT_STARTS = [0, 15, 33, 44];
const CORRECT_OPTIONS = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
const PER_UNIT = 5;
const SHEET_NAME = '個人成績';

// Run once from the Apps Script editor. The spreadsheet remains in your Drive.
function setupScoreReceiver() {
  const properties = PropertiesService.getScriptProperties();
  let sheetId = properties.getProperty('SCORE_SHEET_ID');
  if (!sheetId) {
    const spreadsheet = SpreadsheetApp.create('第6章 板塊運動與岩層的祕密｜個人成績');
    spreadsheet.setSpreadsheetTimeZone('Asia/Taipei');
    const sheet = spreadsheet.getSheets()[0].setName(SHEET_NAME);
    sheet.appendRow(['收件時間', '測驗編號', '版本', '班級', '座號',
      '答對題數', '總題數', '百分制', '6-1 答對', '6-2 答對', '6-3 答對']);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('B:E').setNumberFormat('@');
    sheetId = spreadsheet.getId();
    properties.setProperty('SCORE_SHEET_ID', sheetId);
  }
  let uploadCode = properties.getProperty('SCORE_UPLOAD_CODE');
  if (!uploadCode) {
    uploadCode = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    properties.setProperty('SCORE_UPLOAD_CODE', uploadCode);
  }
  console.log('教師成績表 https://docs.google.com/spreadsheets/d/' + sheetId + '/edit');
  console.log('交給學生的上傳代碼 ' + uploadCode);
}

function doGet() {
  return responsePage_('第6章成績收件端', '請在學生網頁完成測驗，從成績頁上傳。');
}

function doPost(event) {
  const incoming = event && event.parameter || {};
  const properties = PropertiesService.getScriptProperties();
  const sheetId = properties.getProperty('SCORE_SHEET_ID');
  const uploadCode = properties.getProperty('SCORE_UPLOAD_CODE');
  if (!sheetId || !uploadCode) return responsePage_('收件端尚未設定', '請通知老師。');
  if (String(incoming.classCode || '').trim() !== uploadCode) {
    return responsePage_('成績尚未收件', '上傳代碼有誤。請回到測驗頁確認代碼後重送。');
  }

  let record;
  try { record = validateSubmission_(incoming); }
  catch (error) { return responsePage_('成績尚未收件', '班級、座號或測驗資料有誤。請確認後重送。'); }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(SHEET_NAME);
    const size = sheet.getLastRow() - 1;
    const existing = size > 0
      ? sheet.getRange(2, 2, size, 1).createTextFinder(record.id).matchEntireCell(true).findNext()
      : null;
    if (existing) {
      const previous = sheet.getRange(existing.getRow(), 4, 1, 2).getDisplayValues()[0];
      if (previous[0] !== record.classNo || previous[1] !== record.seat) {
        return responsePage_('這份測驗已收件', '若班級或座號需更正，請聯絡老師。');
      }
      const previousScore = sheet.getRange(existing.getRow(), 6, 1, 3).getValues()[0];
      return responsePage_('成績已收到', '此筆成績已收件，無須重傳。' +
        previousScore[2] + ' 分（' + previousScore[0] + '／' + previousScore[1] + ' 題）');
    }
    sheet.appendRow([new Date(), record.id, TEST_VERSION, record.classNo, record.seat,
      record.correct, record.total, record.percent, ...record.byUnit]);
    SpreadsheetApp.flush();
    return responsePage_('成績已收到', record.classNo + ' 班 ' + record.seat + ' 號｜' +
      record.percent + ' 分（' + record.correct + '／' + record.total + ' 題）');
  } catch (error) {
    console.error('Score receiver unavailable');
    return responsePage_('目前無法確認收件', '請回到測驗頁再送一次。同一份測驗只會登記一次。');
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function validateSubmission_(incoming) {
  const classNo = String(incoming.classNo || '');
  const seatInput = String(incoming.seat || '');
  const raw = incoming.attempt;
  if (!/^[0-9]{3}$/.test(classNo) || !/^[0-9]{1,2}$/.test(seatInput) || Number(seatInput) < 1 ||
      typeof raw !== 'string' || raw.length > 1500) throw new Error('Invalid submission');
  const attempt = JSON.parse(raw);
  const total = PER_UNIT * (UNIT_STARTS.length - 1);
  if (!attempt || attempt.version !== TEST_VERSION || typeof attempt.id !== 'string' ||
      !/^[a-f0-9]{32}$/.test(attempt.id) || !Array.isArray(attempt.answers) || attempt.answers.length !== total) {
    throw new Error('Invalid assessment');
  }
  const ids = new Set();
  const counts = [0, 0, 0];
  const byUnit = [0, 0, 0];
  for (const pair of attempt.answers) {
    if (!Array.isArray(pair) || pair.length !== 2) throw new Error('Invalid answer');
    const [id, choice] = pair;
    if (!Number.isInteger(id) || id < 0 || id >= CORRECT_OPTIONS.length || ids.has(id) ||
        !Number.isInteger(choice) || choice < 0 || choice > 3) throw new Error('Invalid answer key');
    ids.add(id);
    const unit = UNIT_STARTS.findIndex((start, index) => id >= start && id < UNIT_STARTS[index + 1]);
    counts[unit]++;
    if (choice === CORRECT_OPTIONS[id]) byUnit[unit]++;
  }
  if (counts.some(count => count !== PER_UNIT)) throw new Error('Invalid selection');
  const correct = byUnit.reduce((sum, count) => sum + count, 0);
  return { id: attempt.id, classNo, seat: String(Number(seatInput)).padStart(2, '0'),
    byUnit, correct, total, percent: Math.round(100 * correct / total) };
}

function responsePage_(title, message) {
  const escape = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"></head>' +
    '<body style="font:18px system-ui,sans-serif;line-height:1.6;background:#faf5ed;color:#252a31">' +
    '<main style="max-width:32rem;margin:10vh auto;padding:1.5rem"><h1>' + escape(title) + '</h1><p>' +
    escape(message) + '</p><p>可關閉此分頁，回到複習網頁。</p></main></body></html>'
  ).setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
