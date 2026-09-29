// 오늘의 일기 도우미 - 저장 + AI 코칭 서버 (학생별 시트 방식)
// 이 시트의 "학생명단" 탭: A=학생(번호/별명), B=탭 번호 또는 파일 ID(자동), C=링크(자동)
// 학생마다 이 파일 안에 탭이 생기고 A=날짜, B=날씨, C=내가 쓴 글, D=AI가 수정해준 글 로 저장됩니다.
// (B열에 긴 파일 ID가 있는 학생은 예전 방식대로 그 별도 파일에 저장합니다)
// 위에서부터 차례로 시도합니다. 모델이 종료되거나 오류가 나면 AI Studio에서 현재 이름을 확인해 여기만 수정
var MODELS = ['gemini-3.5-flash', 'gemini-3.1-flash-lite'];
var ROSTER = '학생명단';
var HEADER = ['날짜', '날씨', '내가 쓴 글', 'AI가 수정해준 글'];

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('일기 도우미').addItem('학생 시트 만들기', 'makeStudentSheets').addToUi();
}

function doPost(e) {
  try {
    var p = PropertiesService.getScriptProperties();
    var d = JSON.parse(e.postData.contents);
    var need = p.getProperty('ACCESS_CODE');
    if (need && d.code !== need) return out({ ok: false, error: '교사 코드가 달라요' });
    if (d.action === 'students') return out({ ok: true, names: readRoster().map(function (x) { return x.name; }) });
    if (d.action === 'coach') return out(coach(d, p));
    if (d.action === 'save') return out(save(d));
    return out({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() { return ContentService.createTextOutput('일기 도우미 연결 OK'); }

function rosterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ROSTER);
  if (!sh) { sh = ss.insertSheet(ROSTER); sh.appendRow(['학생(번호/별명)', '시트ID(자동)', '시트 링크(자동)']); }
  return sh;
}

function readRoster() {
  var sh = rosterSheet(), n = sh.getLastRow(), r = [];
  if (n < 2) return r;
  var v = sh.getRange(2, 1, n - 1, 3).getValues();
  for (var i = 0; i < v.length; i++) {
    var name = String(v[i][0]).trim();
    if (name) r.push({ row: i + 2, name: name, id: String(v[i][1]).trim() });
  }
  return r;
}

// 메뉴: 일기 도우미 > 학생 시트 만들기
// 이 파일 안에 학생마다 탭(시트)을 만듭니다. 이미 만든 학생은 건너뜁니다.
function makeStudentSheets() {
  var main = SpreadsheetApp.getActiveSpreadsheet(), sh = rosterSheet(), list = readRoster();
  var made = 0, skipped = [];
  list.forEach(function (s) {
    if (s.id) return;
    if (s.name === ROSTER || s.name.length > 90 || /[\[\]\*\?:\/\\]/.test(s.name)) { skipped.push(s.name); return; }
    var t = main.getSheetByName(s.name);
    if (!t) { t = main.insertSheet(s.name); t.appendRow(HEADER); }
    sh.getRange(s.row, 2).setValue(t.getSheetId());
    sh.getRange(s.row, 3).setValue(main.getUrl() + '#gid=' + t.getSheetId());
    made++;
  });
  var msg = list.length ? made + '개의 학생 탭을 만들었어요. (이미 있는 학생은 건너뛰었어요)'
                        : '"학생명단" 탭의 A2 칸부터 학생 번호나 별명을 적고 다시 눌러 주세요.';
  if (skipped.length) msg += '\n다음 이름은 탭 이름으로 쓸 수 없어 건너뛰었어요: ' + skipped.join(', ');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function save(d) {
  var list = readRoster(), sh;
  var main = SpreadsheetApp.getActiveSpreadsheet();
  if (list.length) {
    if (!d.student) return { ok: false, error: '내 이름(번호)을 골라 주세요' };
    var s = list.filter(function (x) { return x.name === d.student; })[0];
    if (!s) return { ok: false, error: '명단에 없는 학생이에요' };
    if (!s.id) return { ok: false, error: '학생 시트가 아직 없어요. 선생님께 알려 주세요' };
    if (/^\d+$/.test(s.id)) { // 새 방식: 이 파일 안의 학생 탭 (B열에 탭 번호)
      sh = main.getSheets().filter(function (x) { return x.getSheetId() === Number(s.id); })[0];
      if (!sh) return { ok: false, error: '학생 탭을 찾지 못했어요. 선생님께 알려 주세요' };
    } else { // 예전 방식: 별도 파일 (B열에 파일 ID)
      var other = SpreadsheetApp.openById(s.id);
      sh = other.getSheetByName('일기') || other.insertSheet('일기');
    }
  } else {
    sh = main.getSheetByName('일기') || main.insertSheet('일기'); // 명단이 비어 있으면 이 파일의 일기 탭
  }
  if (sh.getLastRow() === 0) sh.appendRow(HEADER);
  else if (!sh.getRange('D1').getValue()) sh.getRange('D1').setValue(HEADER[3]);
  sh.appendRow([d.date, d.weather, d.mine, d.ai || '']);
  return { ok: true };
}

function coach(d, p) {
  var key = p.getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, error: 'API 키가 설정되지 않았어요' };
  var system =
    '너는 특수교육 교사를 돕는 따뜻한 일기 코치다. 학생이 쓴 일기를 조금 더 풍부하게 다듬는다.\n' +
    '규칙:\n' +
    '1) 학생이 고른 사실(누가·언제·어디서·무엇을·기분)만 사용하고 새로운 사건·인물·장소를 지어내지 않는다.\n' +
    '2) 초등 저학년이 쓸 법한 쉬운 단어와 짧은 문장으로 쓴다. 어른스러운 표현, 한자어, 비유는 쓰지 않는다.\n' +
    '3) 전체 5문장 이내. 학생이 말한 범위 안에서 보이는 것, 들리는 것, 기분 중 한두 가지를 자연스럽게 덧붙인다.\n' +
    '4) 맞춤법과 띄어쓰기를 바로잡는다.\n' +
    '5) 학생의 말투(~다 / ~요)를 그대로 유지한다.\n' +
    '6) feedback은 교사가 학생에게 말하듯 다정한 해요체 2~3문장: 잘한 점 하나를 칭찬하고, 다음에 해 볼 것 하나를 제안한다.\n' +
    '출력은 JSON 한 개만: {"diary":"...","feedback":"..."}';
  var f = d.fields || {};
  var user = '날짜: ' + d.date + '\n날씨: ' + d.weather +
    '\n학생이 고른 내용: 누가=' + (f.who || '') + ', 언제=' + (f.when || '') + ', 어디서=' + (f.where || '') +
    ', 무엇을=' + (f.what || '') + ', 어떻게=' + (f.how || '') + ', 기분=' + (f.feel || '') +
    '\n학생의 일기:\n' + d.draft;
  var body = JSON.stringify({
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: 'user', parts: [{ text: user }] }],
    generationConfig: { responseMimeType: 'application/json', maxOutputTokens: 2048 }
  });
  var j, code, lastErr = 'AI 호출 실패';
  for (var mi = 0; mi < MODELS.length; mi++) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODELS[mi] + ':generateContent';
    for (var t = 0; t < 3; t++) { // 바쁠 때(429/500/503)는 잠깐 쉬었다가 다시 시도
      var res = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-goog-api-key': key }, payload: body
      });
      code = res.getResponseCode();
      j = JSON.parse(res.getContentText());
      if (code === 200) break;
      lastErr = (j.error && j.error.message) || lastErr;
      if (code !== 429 && code !== 500 && code !== 503) break;
      Utilities.sleep(1500 * (t + 1));
    }
    if (code === 200) break; // 성공하면 끝, 실패하면 다음 모델로
  }
  if (code !== 200) return { ok: false, error: lastErr };
  var c = j.candidates && j.candidates[0];
  if (!c || !c.content || !c.content.parts) return { ok: false, error: 'AI가 답을 주지 않았어요 (' + ((c && c.finishReason) || '빈 응답') + ')' };
  var text = c.content.parts.map(function (x) { return x.text || ''; }).join('');
  var m = text.match(/\{[\s\S]*\}/);
  if (!m) return { ok: false, error: 'AI 답변을 읽지 못했어요' };
  var o = JSON.parse(m[0]);
  return { ok: true, diary: o.diary, feedback: o.feedback };
}
