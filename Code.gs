/**
 * 기사 속 왜곡 찾기 — 구글 앱스 스크립트 백엔드
 *
 * 사용법:
 * 1) 구글 스프레드시트를 새로 만든다.
 * 2) 확장 프로그램 > Apps Script 에 들어가서 이 파일 내용을 그대로 붙여넣는다.
 * 3) 저장 후 배포 > 새 배포 > 유형: "웹 앱" 선택
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한: 전체 허용(익명 사용자 포함)
 * 4) 생성된 웹 앱 URL(.../exec 로 끝남)을 복사해서 index.html의 GAS_URL 값에 넣는다.
 * 5) 스프레드시트에 "사례" 탭을 만들고 1행에 헤더를 아래처럼 입력한다:
 *    id | outlet | headline | body | sourceUrl
 *
 *    headline, body 칸에는 문제되는 표현을 아래 형식으로 감싸서 적는다:
 *      [[문제되는 문구|왜 문제인지 설명]]
 *    예: 구로구 모친 살해 [[조현병 20대|질환명과 범죄를 인과관계처럼 연결해 낙인을 강화합니다]], 1심 징역 12년
 *
 *    body 칸에서 문단을 나눌 때는 셀 안에서 Alt+Enter(줄바꿈)를 사용한다.
 *    "참가자" 탭과 "통계" 탭은 스크립트가 자동으로 만든다.
 *
 * 관리자 비밀번호 설정 (필수):
 * 1) Apps Script 편집기 좌측 톱니바퀴(프로젝트 설정) 클릭
 * 2) "스크립트 속성" 항목에서 속성 추가: 키 = ADMIN_TOKEN, 값 = 원하는 비밀번호
 * 3) 저장 후 다시 배포(새 버전으로 배포)
 * 이 비밀번호는 코드나 웹사이트 어디에도 노출되지 않고, 서버에서만 대조됩니다.
 * 사례 추가/수정/삭제는 이제 구글시트가 아니라 웹사이트의 관리자 화면에서 합니다.
 *
 * 즉석 당첨 추첨 설정:
 * 총 당첨자 수와 당첨 확률은 스프레드시트의 "설정" 탭에서 관리한다 (스크립트가 자동 생성).
 * 시트에서 직접 값을 바꿔도 되고, 웹사이트 관리자 화면의 "당첨 추첨 설정"에서 바꿔도 된다 — 둘 다 같은 시트를 본다.
 * 당첨자가 채워지면 그 이후 참여자는 자동으로 "다음 기회에"로 처리된다.
 * 전화번호 기준으로 한 사람은 한 번만 추첨 대상이 되며, 재참여해도 같은 결과만 다시 보여준다.
 * 당첨 현황은 스프레드시트의 "추첨" 탭에서 언제든 확인할 수 있다.
 */

var DEFAULT_MAX_WINNERS = 2;
var DEFAULT_WIN_PROBABILITY = 0.3; // 참여 완료 시 당첨될 확률 (0.3 = 30%). 실제 값은 "설정" 시트 탭에서 관리됩니다.

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'data';
  if (action === 'data') {
    return jsonResponse(getData());
  }
  return jsonResponse({ error: 'unknown action' });
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'invalid body' });
  }

  if (body.action === 'register') {
    registerParticipant(body);
    return jsonResponse({ ok: true });
  }
  if (body.action === 'complete') {
    recordCompletion(body);
    return jsonResponse({ ok: true });
  }
  if (body.action === 'admin_check') {
    return jsonResponse({ ok: checkToken(body.token) });
  }
  if (body.action === 'admin_add') {
    if (!checkToken(body.token)) return unauthorized();
    addCase(body);
    return jsonResponse({ ok: true });
  }
  if (body.action === 'admin_update') {
    if (!checkToken(body.token)) return unauthorized();
    var updated = updateCase(body);
    return jsonResponse({ ok: updated });
  }
  if (body.action === 'admin_delete') {
    if (!checkToken(body.token)) return unauthorized();
    var deleted = deleteCase(body.id);
    return jsonResponse({ ok: deleted });
  }
  if (body.action === 'draw') {
    return jsonResponse(drawWinner(body.phone));
  }
  if (body.action === 'admin_get_settings') {
    if (!checkToken(body.token)) return unauthorized();
    return jsonResponse(getSettings());
  }
  if (body.action === 'admin_save_settings') {
    if (!checkToken(body.token)) return unauthorized();
    saveSettings(body);
    return jsonResponse({ ok: true });
  }
  return jsonResponse({ error: 'unknown action' });
}

function getSettings() {
  var sheet = getSheet('설정');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['key', 'value']);
    sheet.appendRow(['MAX_WINNERS', DEFAULT_MAX_WINNERS]);
    sheet.appendRow(['WIN_PROBABILITY', DEFAULT_WIN_PROBABILITY]);
  }
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[String(data[i][0])] = data[i][1];
  }
  var maxWinners = Number(map['MAX_WINNERS']);
  var winProbability = Number(map['WIN_PROBABILITY']);
  if (isNaN(maxWinners)) maxWinners = DEFAULT_MAX_WINNERS;
  if (isNaN(winProbability)) winProbability = DEFAULT_WIN_PROBABILITY;
  return { maxWinners: maxWinners, winProbability: winProbability };
}

function setSettingValue(sheet, key, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function saveSettings(body) {
  var maxWinners = Math.floor(Number(body.maxWinners));
  var winProbability = Number(body.winProbability);
  if (isNaN(maxWinners) || maxWinners < 0) maxWinners = DEFAULT_MAX_WINNERS;
  if (isNaN(winProbability)) winProbability = DEFAULT_WIN_PROBABILITY;
  winProbability = Math.min(1, Math.max(0, winProbability));

  var sheet = getSheet('설정');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['key', 'value']);
  }
  setSettingValue(sheet, 'MAX_WINNERS', maxWinners);
  setSettingValue(sheet, 'WIN_PROBABILITY', winProbability);
}

function normalizePhone(p) {
  return String(p || '').replace(/[^0-9]/g, '');
}

function drawWinner(phoneRaw) {
  var phone = normalizePhone(phoneRaw);
  if (!phone) return { win: false, error: 'no phone' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet('추첨');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['전화번호', '당첨여부', '일시']);
    }
    var data = sheet.getDataRange().getValues();

    // 이미 추첨에 참여한 전화번호면 이전 결과를 그대로 반환 (재당첨 방지)
    for (var i = 1; i < data.length; i++) {
      if (normalizePhone(data[i][0]) === phone) {
        return { win: data[i][1] === 'Y', already: true };
      }
    }

    var winners = 0;
    for (var j = 1; j < data.length; j++) {
      if (data[j][1] === 'Y') winners++;
    }

    var settings = getSettings();
    var win = false;
    if (winners < settings.maxWinners) {
      win = Math.random() < settings.winProbability;
    }

    sheet.appendRow([phone, win ? 'Y' : 'N', new Date()]);
    return { win: win, already: false };
  } finally {
    lock.releaseLock();
  }
}

function checkToken(token) {
  var secret = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  return !!secret && String(token) === String(secret);
}

function unauthorized() {
  return jsonResponse({ error: 'unauthorized' });
}

function addCase(body) {
  var sheet = getSheet('사례');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'outlet', 'headline', 'body', 'sourceUrl']);
  }
  var id = body.id || ('c' + Date.now());
  sheet.appendRow([id, body.outlet || '', body.headline || '', body.body || '', body.sourceUrl || '']);
}

function updateCase(body) {
  var sheet = getSheet('사례');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.id)) {
      sheet.getRange(i + 1, 2).setValue(body.outlet || '');
      sheet.getRange(i + 1, 3).setValue(body.headline || '');
      sheet.getRange(i + 1, 4).setValue(body.body || '');
      sheet.getRange(i + 1, 5).setValue(body.sourceUrl || '');
      return true;
    }
  }
  return false;
}

function deleteCase(id) {
  var sheet = getSheet('사례');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getData() {
  var caseSheet = getSheet('사례');
  var rows = caseSheet.getDataRange().getValues();
  var cases = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    cases.push({
      id: String(r[0]),
      outlet: r[1] || '',
      headline: r[2] || '',
      body: r[3] || '',
      sourceUrl: r[4] || ''
    });
  }

  var statSheet = getSheet('통계');
  var statRows = statSheet.getDataRange().getValues();
  var stats = {};
  for (var j = 1; j < statRows.length; j++) {
    var sr = statRows[j];
    if (!sr[0]) continue;
    stats[String(sr[0])] = {
      completions: Number(sr[1]) || 0,
      foundSum: Number(sr[2]) || 0,
      problemSum: Number(sr[3]) || 0,
      wrongSum: Number(sr[4]) || 0
    };
  }

  return { cases: cases, stats: stats };
}

function registerParticipant(body) {
  var sheet = getSheet('참가자');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['타임스탬프', '이름', '연락처', '주소', '개인정보동의']);
  }
  sheet.appendRow([
    new Date(),
    body.name || '',
    body.phone || '',
    body.address || '',
    body.consent ? 'Y' : 'N'
  ]);
}

function recordCompletion(body) {
  var id = body.id;
  var foundCount = Number(body.foundCount) || 0;
  var totalProblems = Number(body.totalProblems) || 0;
  var wrongClicks = Number(body.wrongClicks) || 0;

  var sheet = getSheet('통계');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'completions', 'foundSum', 'problemSum', 'wrongSum']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2).setValue((Number(data[i][1]) || 0) + 1);
      sheet.getRange(i + 1, 3).setValue((Number(data[i][2]) || 0) + foundCount);
      sheet.getRange(i + 1, 4).setValue((Number(data[i][3]) || 0) + totalProblems);
      sheet.getRange(i + 1, 5).setValue((Number(data[i][4]) || 0) + wrongClicks);
      return;
    }
  }
  sheet.appendRow([id, 1, foundCount, totalProblems, wrongClicks]);
}
