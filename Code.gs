/**
 * 황산초 통합 도우미 — 오늘의 일정 백엔드
 *
 * doPost : 월간·주간 계획 시트의 "해당 날짜 칸"에 한 줄을 덧붙인다.
 * doGet  : 그 칸의 최신 내용을 읽어준다.
 *          (gviz 는 응답을 수 분간 캐시해서 시트에서 직접 고친 내용이 늦게 보인다.
 *           화면은 gviz 로 빠르게 그린 뒤, 이 doGet 결과로 조용히 교체한다.)
 *
 * 시트 구조: 월별 탭("3월"…"2월") 안에서 [날짜 행] 바로 아래가 [일정 행]이고,
 *            L~R 열(12~18)이 일~토. 한 칸 안에 줄바꿈으로 여러 일정이 쌓인다.
 *            0번 열(A)에는 그 주의 안내 사항이 통째로 들어 있다.
 *
 * ⚠ 배포 방법 (익명 웹앱)
 *  1) Apps Script 편집기 → [서비스 +] → "Google Sheets API" 추가 (식별자: Sheets)
 *     - 익명 웹앱에서는 SpreadsheetApp 같은 내장 서비스가 권한 오류를 낸다.
 *       반드시 고급 서비스(Sheets)를 써야 한다.
 *  2) authorize() 를 한 번 실행해 권한 동의
 *  3) 배포 → 배포 관리 → 기존 배포를 "새 버전"으로 업데이트
 *     (새 배포로 만들면 URL 이 바뀌어 index.html 도 고쳐야 한다)
 */

var SPREADSHEET_ID = '1Z9Jz5F3SMXNHRQIaXrTsGS-B5tiUSKanc2pGdfCrVUY';

var SUN_COL = 12;      // L열 = 일요일 (1-indexed)
var SCAN_RANGE = 40;   // 탭 위에서부터 훑을 행 수
var CACHE_SEC = 20;    // 조회 캐시. 짧게 잡아 준실시간을 유지한다

/** 권한 동의용. 편집기에서 한 번 실행한다. */
function authorize() {
  var meta = Sheets.Spreadsheets.get(SPREADSHEET_ID);
  Logger.log('탭 %s개', meta.sheets.length);
}

/* ────────────────────────────── 조회 ────────────────────────────── */

function doGet(e) {
  try {
    var ymd = String(((e && e.parameter) || {}).date || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) throw new Error('날짜 형식이 올바르지 않습니다.');

    var hit = cacheGet('plan_' + ymd);
    if (hit) return json(hit);

    var loc = locate(ymd);
    var out = {
      success: true,
      date: ymd,
      text: loc.found ? loc.text : '',
      notice: loc.found ? loc.notice : ''
    };
    cachePut('plan_' + ymd, out, CACHE_SEC);
    return json(out);
  } catch (err) {
    return json({ success: false, message: err.message });
  }
}

/* ────────────────────────────── 기록 ────────────────────────────── */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var text = String(body.text || '').trim();
    var ymd  = String(body.date || '').trim();

    if (!text) throw new Error('내용을 입력해주세요.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) throw new Error('날짜 형식이 올바르지 않습니다.');
    if (text.length > 300) throw new Error('내용이 너무 깁니다. (300자 이내)');

    return json(appendPlan(ymd, text));
  } catch (err) {
    return json({ success: false, message: err.message });
  }
}

/** 읽고 → 붙이고 → 쓰는 사이에 다른 편집이 끼어들지 못하게 잠근다 */
function appendPlan(ymd, text) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var loc = locate(ymd);
    if (!loc.found) throw new Error('시트에서 ' + ymd + ' 칸을 찾지 못했습니다.');

    var line = '-' + text;
    var after = loc.text.trim() ? loc.text.replace(/\s+$/, '') + '\n' + line : line;

    Sheets.Spreadsheets.Values.update(
      { values: [[after]] },
      SPREADSHEET_ID, loc.cell,
      { valueInputOption: 'USER_ENTERED' }
    );

    // 방금 바꾼 날짜의 조회 캐시를 지워 다음 조회가 최신을 보게 한다
    try { CacheService.getScriptCache().remove('plan_' + ymd); } catch (ignore) {}

    return { success: true, cell: loc.cell, added: line };
  } finally {
    lock.releaseLock();
  }
}

/* ──────────────────────── 칸 위치 찾기 (공용) ──────────────────────── */

/**
 * 해당 날짜의 일정 칸 위치와 현재 내용을 찾는다.
 * @return {found, cell, text, notice}
 */
function locate(ymd) {
  var parts = ymd.split('-');
  var target = new Date(+parts[0], +parts[1] - 1, +parts[2]);

  // 탭은 "9월"처럼 월 이름으로 찾는다. gid 로 찾으면 누가 탭을 지우고
  // 다시 만들 때마다 여기와 index.html 을 고쳐야 했다. (2026-09-02 실제 발생)
  var title = (target.getMonth() + 1) + '월';

  // 날짜를 시리얼 숫자로 받아야 '일 숫자'와 확실히 구분된다
  var res = Sheets.Spreadsheets.Values.get(SPREADSHEET_ID, "'" + title + "'!A1:R" + SCAN_RANGE, {
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER'
  });
  var rows = res.values || [];

  var col = SUN_COL + target.getDay();
  var dateRow = findDateRow(rows, target, col);
  if (dateRow < 0) return { found: false };

  var planRow = dateRow + 1;                       // 날짜 행 바로 아래
  var row = rows[planRow] || [];
  return {
    found: true,
    cell: "'" + title + "'!" + colLetter(col) + (planRow + 1),
    text: String(row[col - 1] || ''),
    notice: String(row[0] || '')                   // A열 = 그 주 안내 사항
  };
}

/**
 * 날짜 행 찾기.
 * 날짜 셀은 서식에 따라 시리얼 숫자(45000+)이거나 그냥 일(1~31) 숫자다.
 * 방학 달 탭은 7칸이 전부 일 숫자라 시리얼이 하나도 없다.
 */
function findDateRow(rows, target, col) {
  var wantSerial = toSerial(target);
  var wantDay = target.getDate();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || [];
    var dateLike = 0;
    for (var j = SUN_COL; j <= SUN_COL + 6; j++) {
      var v = row[j - 1];
      if (typeof v === 'number' && (v > 1000 || (v >= 1 && v <= 31))) dateLike++;
    }
    if (dateLike < 4) continue;   // 날짜 행이 아니다

    var cell = row[col - 1];
    if (typeof cell !== 'number') continue;
    if (cell > 1000 ? cell === wantSerial : cell === wantDay) return i;
  }
  return -1;
}

/* ────────────────────────────── 유틸 ────────────────────────────── */

/** 1899-12-30 을 0 으로 하는 스프레드시트 날짜 시리얼 */
function toSerial(date) {
  var utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round((utc - Date.UTC(1899, 11, 30)) / 86400000);
}

function colLetter(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cacheGet(key) {
  try {
    var v = CacheService.getScriptCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch (e) { return null; }
}

function cachePut(key, obj, sec) {
  try {
    var s = JSON.stringify(obj);
    if (s.length < 95000) CacheService.getScriptCache().put(key, s, sec);
  } catch (e) {}
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ──────────────────── 자체 점검 (시트를 건드리지 않는다) ──────────────────── */

function selfCheck() {
  var tests = ['2026-07-20', '2026-08-12', '2026-04-15'];
  tests.forEach(function (ymd) {
    var loc = locate(ymd);
    if (!loc.found) throw new Error(ymd + ' 칸을 찾지 못했습니다.');
    Logger.log('%s → %s | 현재 %s줄', ymd, loc.cell,
      loc.text ? loc.text.split('\n').length : 0);
  });
  Logger.log('자체 점검 통과');
}
