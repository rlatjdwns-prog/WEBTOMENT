/**
 * 미디어 왜곡 토너먼트 — 구글 앱스 스크립트 백엔드
 *
 * 사용법:
 * 1) 구글 스프레드시트를 새로 만든다.
 * 2) 확장 프로그램 > Apps Script 에 들어가서 이 파일 내용을 그대로 붙여넣는다.
 * 3) 저장 후 배포 > 새 배포 > 유형: "웹 앱" 선택
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한: 전체 허용(익명 사용자 포함)
 * 4) 생성된 웹 앱 URL(.../exec 로 끝남)을 복사해서 index.html의 GAS_URL 값에 넣는다.
 * 5) 스프레드시트에 "사례" 탭을 만들고 1행에 헤더를 아래처럼 입력한다:
 *    id | outlet | title | quote | description | sourceUrl
 *    2행부터 실제 보도 왜곡 사례를 한 줄씩 입력하면 된다. (id는 c1, c2 처럼 겹치지 않게)
 *    "참가자" 탭과 "통계" 탭은 스크립트가 자동으로 만든다.
 */

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
  if (body.action === 'vote') {
    recordVote(body.winnerId, body.loserId);
    return jsonResponse({ ok: true });
  }
  return jsonResponse({ error: 'unknown action' });
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
      title: r[2] || '',
      quote: r[3] || '',
      description: r[4] || '',
      sourceUrl: r[5] || ''
    });
  }

  var statSheet = getSheet('통계');
  var statRows = statSheet.getDataRange().getValues();
  var stats = {};
  for (var j = 1; j < statRows.length; j++) {
    var sr = statRows[j];
    if (!sr[0]) continue;
    stats[String(sr[0])] = {
      appearances: Number(sr[1]) || 0,
      wins: Number(sr[2]) || 0
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

function recordVote(winnerId, loserId) {
  var sheet = getSheet('통계');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'appearances', 'wins']);
  }
  if (winnerId) bumpStat(sheet, winnerId, true);
  if (loserId) bumpStat(sheet, loserId, false);
}

function bumpStat(sheet, id, won) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var appearances = (Number(data[i][1]) || 0) + 1;
      var wins = (Number(data[i][2]) || 0) + (won ? 1 : 0);
      sheet.getRange(i + 1, 2).setValue(appearances);
      sheet.getRange(i + 1, 3).setValue(wins);
      return;
    }
  }
  sheet.appendRow([id, 1, won ? 1 : 0]);
}
