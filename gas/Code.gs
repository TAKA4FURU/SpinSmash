// SpinSmash オンラインランキング（Google Apps Script）
// スプレッドシートの「拡張機能 → Apps Script」にこのコードを貼り付け、ウェブアプリとして公開してください。
// 記録はこのスプレッドシートの「scores」シートに1行ずつ保存されます。

const SHEET_NAME = 'scores';
const TOP_N = 50;          // 返す件数
const MAX_SCORE = 1e12;      // これを超えるスコアは不正として受け付けない（ボールが大量に増えるので大きめに）

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['name', 'score', 'stage', 'rally', 'date']);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function topScores_() {
  const rows = getSheet_().getDataRange().getValues().slice(1)
    .filter(r => typeof r[1] === 'number')
    .sort((a, b) => b[1] - a[1]);
  return rows;
}

// ランキングの取得（GET）
function doGet() {
  const top = topScores_().slice(0, TOP_N).map(r => ({
    name: String(r[0]), score: r[1], stage: r[2], rally: r[3],
    date: r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Tokyo', 'M/d') : String(r[4]),
  }));
  return json_({ ok: true, top });
}

// 記録の送信（POST）
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  const name = String(d.name || '').replace(/[<>\u0000-\u001f]/g, '').trim().slice(0, 10) || 'NONAME';
  const score = Math.floor(Number(d.score));
  const stage = Math.floor(Number(d.stage));
  const rally = Math.floor(Number(d.rally)) || 0;
  if (!(score >= 0 && score <= MAX_SCORE) || !(stage >= 1 && stage <= 10)) {
    return json_({ ok: false, error: 'invalid' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet_().appendRow([name, score, stage, rally, new Date()]);
  } finally {
    lock.releaseLock();
  }
  // 順位（同点は上位扱い）
  const rank = topScores_().filter(r => r[1] > score).length + 1;
  return json_({ ok: true, rank });
}
