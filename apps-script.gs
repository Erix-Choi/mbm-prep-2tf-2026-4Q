/* MBM 세션 준비표 — 공용 상태 저장소 (구글 시트 + Apps Script 웹 앱)
   1) 아래 PIN을 팀 PIN으로 바꾸고 저장  2) 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자)  3) 나온 /exec 주소를 data.js의 api에 넣기 */
const PIN = '0000'; // ← 팀 PIN으로 바꾸기 (이 파일 사본에만, 저장소에는 올리지 않기)

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('status');
  if (!sh) {
    sh = ss.insertSheet('status');
    sh.appendRow(['ts', 'view', 'id', 'state', 'pct', 'date', 'who', 'note', 'when (KST)']);
    sh.getRange('F:F').setNumberFormat('@');
  }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  const v = sheet_().getDataRange().getValues().slice(1);
  return out_({ ok: true, rows: v.map(r => ({ ts: r[0], view: r[1], id: r[2], state: r[3], pct: r[4], date: String(r[5]), who: r[6], note: r[7] })) });
}
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad request' }); }
  if (String(d.pin) !== PIN) return out_({ ok: false, error: 'PIN' });
  if (['todo', 'doing', 'done'].indexOf(d.state) < 0 || !d.view || !d.id) return out_({ ok: false, error: 'bad data' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    sheet_().appendRow([now.getTime(), String(d.view), String(d.id), d.state, Number(d.pct) || 0, String(d.date || ''),
      String(d.who || '').slice(0, 20), String(d.note || '').slice(0, 200), Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm')]);
  } finally { lock.releaseLock(); }
  return out_({ ok: true });
}
