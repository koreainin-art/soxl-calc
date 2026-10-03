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
  const p = e.parameter || {};
  if (p.token !== TOKEN) return out_({ ok: false, err: '비밀번호가 맞지 않습니다' });
  if (p.op === 'close') return out_({ ok: true, closes: closes_() });
  return out_({ ok: true, rows: readAll_(sheet_()) });
}

// SOXL 최근 종가 (앱의 전일종가 자동 입력용). 장중 미완성 봉은 제외, 10분 캐시
function closes_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('soxl-closes');
  if (hit) return JSON.parse(hit);
  let rows = [];
  try {
    const r = UrlFetchApp.fetch('https://query1.finance.yahoo.com/v8/finance/chart/SOXL?range=3mo&interval=1d',
      { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (r.getResponseCode() === 200) {
      const j = JSON.parse(r.getContentText()).chart.result[0];
      const ts = j.timestamp || [], cl = j.indicators.quote[0].close || [];
      for (let i = 0; i < ts.length; i++) if (cl[i] != null)
        rows.push([Utilities.formatDate(new Date(ts[i] * 1000), 'America/New_York', 'yyyy-MM-dd'), Math.round(cl[i] * 100) / 100]);
    }
  } catch (x) {}
  if (!rows.length) {
    try {
      const t = UrlFetchApp.fetch('https://stooq.com/q/d/l/?s=soxl.us&i=d', { muteHttpExceptions: true }).getContentText();
      t.trim().split('\n').slice(1).slice(-70).forEach(l => { const c = l.split(','); if (c.length >= 5 && +c[4] > 0) rows.push([c[0], Math.round(+c[4] * 100) / 100]); });
    } catch (x) {}
  }
  // 오늘(미국 날짜) 봉은 장 마감(16:15 ET) 이후에만 확정 종가로 사용
  const now = new Date(), today = Utilities.formatDate(now, 'America/New_York', 'yyyy-MM-dd');
  const hm = +Utilities.formatDate(now, 'America/New_York', 'HHmm');
  if (rows.length && rows[rows.length - 1][0] === today && hm < 1615) rows.pop();
  if (rows.length) cache.put('soxl-closes', JSON.stringify(rows), 600);
  return rows;
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

// 처음 한 번 실행: 외부 데이터(종가) 가져오기 권한 승인 + 동작 확인
function 종가확인() { Logger.log(JSON.stringify(closes_().slice(-5))); }
function 설정확인() { sheet_(); Logger.log('기록 시트 준비 완료'); }
