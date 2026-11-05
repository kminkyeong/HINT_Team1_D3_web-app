/**
 * 통합 업무 관리 웹앱 - 라우팅(doGet) / 초기 설정 / 공통 함수
 * 화면 파일: IndexHome, IndexMeeting, IndexBudget, IndexMinutes + 공통 메뉴 SharedUI
 *
 * ⚠ 보안 규칙
 *   Apps Script에서는 이름이 '_'로 끝나지 않는 함수를 웹 화면(google.script.run)에서
 *   누구나 호출할 수 있습니다.
 *   - 화면에서 쓰는 함수  : 첫 줄에서 assertAccess_() 로 접근 권한 확인
 *   - 관리용 함수(설정·초기화·테스트) : 첫 줄에서 assertOwner_() 로 보호
 *   - 내부 도우미 함수    : 이름 끝에 '_' 를 붙여 외부 호출 차단
 */

/* ===================== 공통: 스프레드시트 / 시간대 / URL ===================== */

function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('스프레드시트에 연결된 Apps Script에서 setup을 먼저 실행하세요.');
  return SpreadsheetApp.openById(id);
}

/** 모든 모듈이 같은 시간대를 쓰도록 통일 (프로젝트 설정에서 Asia/Seoul 지정) */
function appTz_() {
  return Session.getScriptTimeZone() || 'Asia/Seoul';
}

/** 웹앱 주소. 스크립트 속성 WEB_APP_URL(…/exec)이 있으면 우선 사용 */
function appBase_() {
  const fixed = String(PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '').trim();
  if (fixed.indexOf('https://') === 0) return fixed;
  return ScriptApp.getService().getUrl() || '';
}

/* ===================== 공통: Gemini API (OCR · STT · 요약 모두 담당) ===================== */

/**
 * Google AI Studio(aistudio.google.com)에서 무료로 발급한 Gemini API 키
 * 스크립트 속성 GEMINI_API_KEY 에 등록합니다. (결제 연결 없이 무료 등급으로 사용)
 */
function geminiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('프로젝트 설정 → 스크립트 속성에 GEMINI_API_KEY를 추가하세요. (aistudio.google.com에서 무료 발급)');
  return key.trim();
}

/**
 * 무료로 쓸 수 있는 Flash 계열 후보 (앞에서부터 시도)
 * Gemini 모델은 자주 바뀌고, 키(계정)마다 쓸 수 있는 모델도 다릅니다.
 * 그래서 이름을 하나로 고정하지 않고, 이 키로 '실제로 호출되는' 모델을 찾아 GEMINI_MODEL에 저장합니다.
 */
const GEMINI_CANDIDATES_ = [
  'gemini-3.6-flash', 'gemini-3.8-flash', 'gemini-3-flash-preview',
  'gemini-3.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-2.5-flash'
];
const GEMINI_BASE_ = 'https://generativelanguage.googleapis.com/v1beta/models';

/** 사용할 모델: 스크립트 속성 GEMINI_MODEL, 없으면 자동으로 찾아서 저장 */
function geminiModel_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || detectGeminiModel_(false);
}

/**
 * Gemini generateContent 를 한 번 호출 (오류를 던지지 않고 결과만 돌려줌)
 * @return {{ok: boolean, code: number, message: string, json: Object}}
 */
function callGemini_(model, parts, opt) {
  opt = opt || {};
  const key = geminiKey_();
  const body = {
    contents: [{ role: 'user', parts: parts }],
    generationConfig: { temperature: 0.2 }
  };
  if (opt.system) body.systemInstruction = { parts: [{ text: opt.system }] };
  if (opt.schema) {
    body.generationConfig.responseMimeType = 'application/json';
    body.generationConfig.responseSchema = opt.schema;
  }
  const res = UrlFetchApp.fetch(GEMINI_BASE_ + '/' + encodeURIComponent(model) + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key }, // 키를 URL이 아닌 헤더로 보내 로그에 남지 않게 함
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  let json = {};
  try { json = JSON.parse(res.getContentText()); } catch (e) {}
  const message = String((json.error && json.error.message) || '').split(key).join('[숨김]');
  return { ok: code === 200, code, message, json };
}

/**
 * 이 키로 호출 가능한 모델을 찾아 스크립트 속성 GEMINI_MODEL 에 저장하고 이름을 돌려줌
 * 1) 모델 목록(models.list)에서 이 키가 쓸 수 있는 모델 확인
 * 2) 후보를 차례로 짧게 호출해 보고, 처음 성공한 모델을 저장
 *    (무료 등급이 없는 모델은 429, 사라진 모델은 404 → 다음 후보로)
 */
function detectGeminiModel_(log) {
  const props = PropertiesService.getScriptProperties();
  const key = geminiKey_();
  const say = msg => { if (log) Logger.log(msg); };

  let available = null;
  const res = UrlFetchApp.fetch(GEMINI_BASE_ + '?pageSize=1000', {
    headers: { 'x-goog-api-key': key }, muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) {
    try {
      available = (JSON.parse(res.getContentText()).models || [])
        .filter(m => (m.supportedGenerationMethods || []).includes('generateContent'))
        .map(m => String(m.name).replace(/^models\//, ''));
    } catch (e) { available = null; }
  } else if (/API key/i.test(res.getContentText())) {
    throw new Error('Gemini API 키가 올바르지 않습니다. 스크립트 속성 GEMINI_API_KEY를 확인하세요.');
  }

  // 사용자가 직접 넣은 모델 → 우선 후보 → 목록에 있는 다른 Flash 모델(최신순) 순서로 시도
  const manual = props.getProperty('GEMINI_MODEL');
  let candidates = [manual].concat(GEMINI_CANDIDATES_).filter(Boolean);
  if (available) {
    const extra = available
      .filter(n => /flash/.test(n) && !/(image|tts|live|audio|embedding|computer|robotics)/.test(n))
      .sort().reverse();
    candidates = candidates.filter(n => available.includes(n)).concat(extra);
    say('이 키로 보이는 Flash 모델: ' + (extra.join(', ') || '(없음)'));
  }
  candidates = [...new Set(candidates)];

  for (const model of candidates) {
    const r = callGemini_(model, [{ text: '"정상"이라고만 답하세요.' }]);
    if (r.ok) {
      props.setProperty('GEMINI_MODEL', model);
      say('✅ 사용 가능: ' + model + ' → 스크립트 속성 GEMINI_MODEL 에 저장했습니다.');
      return model;
    }
    say('  ✗ ' + model + ' : HTTP ' + r.code + ' ' + r.message.slice(0, 120));
    if (r.code === 400 && /API key/i.test(r.message)) {
      throw new Error('Gemini API 키가 올바르지 않습니다. 스크립트 속성 GEMINI_API_KEY를 확인하세요.');
    }
  }
  throw new Error('이 API 키로 사용할 수 있는 무료 Gemini 모델을 찾지 못했습니다. ' +
    'AI Studio(aistudio.google.com)에서 사용 가능한 모델명을 확인해 스크립트 속성 GEMINI_MODEL에 넣거나, 잠시 뒤 다시 시도하세요.');
}

/**
 * Gemini 호출 → 응답 텍스트 반환 (OCR·전사·요약이 모두 이 함수를 사용)
 * @param {Object[]} parts 보낼 내용. 예) [{text: '질문'}, {inlineData: {mimeType: 'image/jpeg', data: base64}}]
 * @param {{system?: string, schema?: Object}} opt
 *        system: 역할·규칙 지시문 / schema: 이 형식의 JSON으로만 답하게 할 때 (responseSchema)
 */
function gemini_(parts, opt) {
  let model = geminiModel_();
  let r = callGemini_(model, parts, opt);

  // 저장된 모델이 사라졌으면(404) 한 번만 새 모델을 찾아 다시 시도
  if (r.code === 404) {
    PropertiesService.getScriptProperties().deleteProperty('GEMINI_MODEL');
    model = detectGeminiModel_(false);
    r = callGemini_(model, parts, opt);
  }

  if (!r.ok) {
    if (r.code === 429) throw new Error('Gemini 무료 사용량 한도에 도달했습니다. 1분 뒤 다시 시도하고, 계속되면 하루 한도이므로 내일 다시 시도하세요.');
    if (/API key/i.test(r.message)) throw new Error('Gemini API 키가 올바르지 않습니다. 스크립트 속성 GEMINI_API_KEY를 확인하세요.');
    if (r.code === 404) throw new Error(`Gemini 모델(${model})을 찾을 수 없습니다. 편집기에서 testGeminiApiKey를 실행해 사용 가능한 모델을 다시 찾으세요.`);
    throw new Error('Gemini HTTP ' + r.code + ': ' + (r.message || '요청 실패').slice(0, 300));
  }

  const json = r.json;
  if (json.promptFeedback && json.promptFeedback.blockReason) {
    throw new Error('Gemini가 이 요청을 처리하지 않았습니다. (' + json.promptFeedback.blockReason + ')');
  }
  const cand = (json.candidates || [])[0] || {};
  const text = ((cand.content && cand.content.parts) || [])
    .filter(p => !p.thought).map(p => p.text || '').join('').trim();
  if (cand.finishReason === 'MAX_TOKENS') throw new Error('응답이 너무 길어 잘렸습니다. 더 짧은 음성으로 다시 시도하세요.');
  if (!text) throw new Error('Gemini 응답이 비어 있습니다.' + (cand.finishReason ? ' (' + cand.finishReason + ')' : ''));
  return text;
}

/**
 * 관리자용: 편집기에서 실행 → 이 키로 쓸 수 있는 무료 모델을 찾아 저장하고 연결을 점검 (실행 로그 확인)
 * 모델 관련 오류(404 등)가 나면 언제든 다시 실행하면 됩니다.
 */
function testGeminiApiKey() {
  assertOwner_();
  const model = detectGeminiModel_(true);
  const answer = gemini_([{ text: '연결 테스트입니다. "정상"이라고만 답하세요.' }]);
  Logger.log('✅ Gemini 연결 정상 · 모델: ' + model + ' · 응답: ' + answer);
  return answer;
}

function escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function safeToken_(s) {
  return String(s || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 64);
}

/* ===================== 공통: 권한 ===================== */

function currentEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; }
}

function ownerEmail_() {
  try { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; }
}

/**
 * 관리용 함수 보호.
 * 편집기·스프레드시트 메뉴에서 실행하면 '실행 사용자 = 로그인 사용자'라서 통과하고,
 * 웹앱(실행 사용자: 나)에서 다른 사람이 호출하면 두 값이 달라 차단됩니다.
 */
function assertOwner_() {
  const me = currentEmail_();
  if (!me || me !== ownerEmail_()) {
    throw new Error('관리자 전용 기능입니다. 스프레드시트 메뉴나 Apps Script 편집기에서 실행하세요.');
  }
}

/**
 * 웹앱 사용자 제한 (스크립트 속성)
 *  - ALLOWED_DOMAIN : 회사 도메인 (예: company.com)
 *  - ALLOWED_EMAILS : 허용할 이메일 목록 (쉼표 구분)
 * 둘 다 비어 있으면 배포 설정의 '액세스 권한'에만 의존합니다.
 * ※ 같은 Google Workspace 도메인 안에서만 사용자 이메일을 확인할 수 있습니다.
 */
function accessAllowed_() {
  const props = PropertiesService.getScriptProperties();
  const domain = String(props.getProperty('ALLOWED_DOMAIN') || '').trim().toLowerCase().replace(/^@/, '');
  const list = String(props.getProperty('ALLOWED_EMAILS') || '').split(',')
    .map(s => s.trim().toLowerCase()).filter(Boolean);
  if (!domain && !list.length) return true;
  const me = currentEmail_();
  if (!me) return false;
  if (me === ownerEmail_() || list.includes(me)) return true;
  return !!domain && me.endsWith('@' + domain);
}

function assertAccess_() {
  if (!accessAllowed_()) {
    throw new Error('이 기능을 사용할 권한이 없습니다. 회사 계정으로 로그인했는지 확인하세요.');
  }
}

/* ===================== 화면 생성 ===================== */

function sharedNav_(active) {
  return HtmlService.createHtmlOutputFromFile('SharedUI').getContent()
    .replace(/__ACTIVE__/g, active);
}

/**
 * HTML 파일을 읽어 자리표시자를 채우고, active가 있으면 공통 메뉴(SharedUI)를 <body> 바로 뒤에 넣습니다.
 * @param {string} file  HTML 파일 이름
 * @param {string} active 메뉴 강조 키(home/meeting/budget/minutes). 빈 값이면 메뉴 없음(대화상자용)
 * @param {Object} replacements { '__자리표시자__': 값 }
 */
function page_(file, active, replacements) {
  let html = HtmlService.createHtmlOutputFromFile(file).getContent();
  Object.keys(replacements || {}).forEach(k => { html = html.split(k).join(replacements[k]); });
  if (active) {
    const nav = sharedNav_(active);
    html = html.replace(/<body([^>]*)>/i, m => m + nav);
  }
  html = html.split('__APP_BASE__').join(appBase_());
  return HtmlService.createHtmlOutput(html)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function denied_() {
  const me = currentEmail_();
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;max-width:560px;margin:60px auto;padding:24px;border:1px solid #dfe5ec">' +
    '<h2 style="color:#002c5f">접근 권한이 없습니다</h2>' +
    '<p>회사 계정으로 로그인한 뒤 다시 접속하세요.</p>' +
    '<p style="color:#5f6f84">현재 계정: ' + escHtml_(me || '확인 불가') + '</p></div>'
  ).setTitle('접근 권한 없음');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!accessAllowed_()) return denied_();

  // 참석자 개인 응답 링크 (?r=<토큰>)
  if (p.r) return render_('respond', { token: p.r }, true).setTitle('회의 참석 응답');
  // 관리자 링크 (?m=<회의ID>&admin=<관리키>)
  if (p.m) {
    return render_(p.admin ? 'manage' : 'view', { mid: p.m, key: p.admin }, true).setTitle('회의 관리');
  }
  if (p.page === 'create') return render_('create', {}, true).setTitle('새 회의 등록');
  if (p.page === 'budget') return page_('IndexBudget', 'budget').setTitle('팀 예산 및 영수증');
  if (p.page === 'minutes') {
    return page_('IndexMinutes', 'minutes', { '__MEETING_ID__': safeToken_(p.meeting) })
      .setTitle('스마트 회의록 · AI 요약');
  }
  return page_('IndexHome', 'home').setTitle('통합 업무 관리');
}

/* ===================== 초기 설정 / 메뉴 ===================== */

function setup() {
  assertOwner_();
  setupMeeting();
  setupBudget();
  setupMinutes();
  return '회의 일정, 예산, AI 회의록 시트 초기화 완료';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂 통합 업무 관리')
    .addItem('① 전체 시트 초기 설정', 'setup')
    .addItem('② 새 회의 등록', 'showCreateDialog')
    .addItem('회의 상태 지금 갱신', 'updateMeetingStatuses')
    .addItem('③ 회의 시트만 설정', 'setupMeeting')
    .addItem('④ 예산 시트만 설정', 'setupBudget')
    .addItem('⑤ 회의록 시트만 설정', 'setupMinutes')
    .addSeparator()
    .addItem('⚠ 예산·지출 데이터 초기화', 'resetData')
    .addToUi();
}
