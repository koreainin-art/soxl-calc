// SOXL 매매기록 — 구글 시트 백업용 Apps Script
// 1) 아래 TOKEN 을 본인만 아는 비밀번호로 바꾸세요 (앱에 똑같이 입력).
const TOKEN = '여기에-비밀번호';

const HEAD = ['날짜','사이클차수','사이클시드','잔금','매수예약가1','매수예약1개수','매수예약가2','매수예약2개수','매도예약가1','매도예약1개수','매도예약가2','매도예약가2개수','매도예약가3','매도예약가3개수','전일종가','전일평단','전일매수체결가','전일매수개수','전일매도체결가','전일매도개수','누적보유수','당일실현','MA20'];
const QTY = [5,7,9,11,13,17,19,20];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('기록');
  if (!sh) sh = ss.insertSheet('기록');
  if (sh.getLastRow() === 0) { sh.appendRow(HEAD); sh.setFrozenRows(1); sh.setFrozenColumns(1); }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function key_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const g = String(v || '').replace(/\D/g, '');
  return g.length === 8 ? g.slice(0,4) + '-' + g.slice(4,6) + '-' + g.slice(6) : '';
}
function readAll_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEAD.length).getValues()
    .map(r => { const k = key_(r[0]); if (!k) return null;
      const row = [k, Number(String(r[1] || '').replace(/\D/g, '')) || 0];
      for (let j = 2; j < HEAD.length; j++) row.push(Number(r[j]) || 0);
      return row; })
    .filter(Boolean);
}
function writeAll_(sh, rows) {
  rows.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
  const n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, HEAD.length).clearContent();
  if (!rows.length) return;
  const vals = rows.map(r => { const g = r[0].split('-'); const x = r.slice();
    x[0] = new Date(+g[0], +g[1] - 1, +g[2]); x[1] = x[1] ? x[1] + '차' : ''; return x; });
  sh.getRange(2, 1, vals.length, HEAD.length).setValues(vals);
  sh.getRange(2, 1, vals.length, 1).setNumberFormat('yyyy.mm.dd');
  for (let j = 2; j < HEAD.length; j++)
    sh.getRange(2, j + 1, vals.length, 1).setNumberFormat(QTY.indexOf(j) >= 0 ? '#,##0' : (j === 15 ? '#,##0.000' : '#,##0.00'));
}

// 앱에서 전체 기록 불러오기
function doGet(e) {
  if ((e.parameter || {}).token !== TOKEN) return out_({ ok: false, err: '비밀번호가 맞지 않습니다' });
  return out_({ ok: true, rows: readAll_(sheet_()) });
}

// 앱에서 기록 저장(put) / 삭제(del)
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: '형식 오류' }); }
  if (b.token !== TOKEN) return out_({ ok: false, err: '비밀번호가 맞지 않습니다' });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    const map = {};
    (b.replace ? [] : readAll_(sh)).forEach(r => map[r[0]] = r);
    if (b.op === 'put') (b.rows || []).forEach(r => { if (r && r[0]) map[r[0]] = r; });
    if (b.op === 'del') (b.dates || []).forEach(d => delete map[d]);
    const rows = Object.keys(map).map(k => map[k]);
    writeAll_(sh, rows);
    return out_({ ok: true, count: rows.length });
  } finally { lock.releaseLock(); }
}
