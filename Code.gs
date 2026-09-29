// 오늘의 일기 도우미 - 저장 + AI 코칭 서버
// 시트 "일기": A=날짜, B=날씨, C=내가 쓴 글, D=AI가 수정해준 글
var MODEL = 'gemini-3.5-flash'; // 모델이 종료되거나 오류가 나면 AI Studio에서 현재 이름을 확인해 여기만 수정

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var p = PropertiesService.getScriptProperties();
    var d = JSON.parse(e.postData.contents);
    var need = p.getProperty('ACCESS_CODE');
    if (need && d.code !== need) return out({ ok: false, error: '교사 코드가 달라요' });
    if (d.action === 'coach') return out(coach(d, p));
    if (d.action === 'save') return out(save(d));
    return out({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() { return ContentService.createTextOutput('일기 도우미 연결 OK'); }

function save(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('일기') || ss.insertSheet('일기');
  if (sh.getLastRow() === 0) sh.appendRow(['날짜', '날씨', '내가 쓴 글', 'AI가 수정해준 글']);
  else if (!sh.getRange('D1').getValue()) sh.getRange('D1').setValue('AI가 수정해준 글'); // 예전 시트 보완
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
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL + ':generateContent';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: [{ role: 'user', parts: [{ text: user }] }],
      generationConfig: { responseMimeType: 'application/json', maxOutputTokens: 2048 }
    })
  });
  var j = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) return { ok: false, error: (j.error && j.error.message) || 'AI 호출 실패' };
  var c = j.candidates && j.candidates[0];
  if (!c || !c.content || !c.content.parts) return { ok: false, error: 'AI가 답을 주지 않았어요 (' + ((c && c.finishReason) || '빈 응답') + ')' };
  var text = c.content.parts.map(function (x) { return x.text || ''; }).join('');
  var m = text.match(/\{[\s\S]*\}/);
  if (!m) return { ok: false, error: 'AI 답변을 읽지 못했어요' };
  var o = JSON.parse(m[0]);
  return { ok: true, diary: o.diary, feedback: o.feedback };
}
